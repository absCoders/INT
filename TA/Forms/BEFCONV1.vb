Public Class BEFCONV1
    Dim tblBETCONVS As DataTable
    Dim sqlBETCONVS As String = "Select * from MYS_SCHEMA_NAME.MYS_TABLE_NAME"
    Dim CONSTR As String = "server=192.168.96.143;user=root;database=actemra;port=3306;password=office;allow zero datetime=yes;"
    Dim rowBETCONV1 As DataRow
    Dim mySQL_Schemas As New List(Of String)
    Dim ORA_TBL_SCHEMA As New Dictionary(Of String, Integer)
    Dim ALL_SCHEMAs As New Dictionary(Of String, Dictionary(Of String, Integer))
    Dim BeforeRowsDeleted As New List(Of String)
    Dim SCHEMAs As New Dictionary(Of String, Integer)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "BETCONV1", "*", 1)

            ASCMAIN1.sql = "Select BETCONV1.*" _
                & " from BETCONV1"
            Create_TDA(.Tables.Add, "BETCONVX", "**", 0, False, "", 1)
            With .Tables("BETCONVX")
                .Columns.Add("LAST_CONV", GetType(System.DateTime))
                .Columns.Add("SECS", GetType(System.Int32))
                .Columns.Add("ROW_COUNT", GetType(System.Int32))
                .Columns.Add("SQL_ERROR")
            End With

            ASCMAIN1.sql = "Select * from BETCONV2 where ORA_TABLE_NAME = :PARM1"
            Create_TDA(.Tables.Add, "BETCONV2", "**", 0, True, "V", 4)

            ASCMAIN1.sql = "Select * from BETCONV3 where ORA_TABLE_NAME = :PARM1"
            Create_TDA(.Tables.Add, "BETCONV3", "**", 0, True, "V", 3)
            Create_Relation("BETCONVX", "BETCONV3", "ORA_TABLE_NAME")
            With .Tables("BETCONV3")
                .Columns.Add("LAST_CONV", GetType(System.DateTime))
                .Columns.Add("SECS", GetType(System.Int32))
                .Columns.Add("ROW_COUNT", GetType(System.Int32))
                .Columns.Add("SQL_ERROR")
                .Columns.Add("SQL_TEST")
                For C As Integer = 0 To 199
                    .Columns.Add("COL_" & Format(C, "000"))
                Next
            End With

            ASCMAIN1.sql = "Select BETCONV0.*" _
                & " from BETCONV0"
            Create_TDA(.Tables.Add, "BETCONV0", "**", 0, True, "", 2)

            Create_Relation("BETCONV0", "BETCONV3", "SCHEMA_NAME,TABLE_NAME", "MYS_SCHEMA_NAME,MYS_TABLE_NAME")
            With .Tables("BETCONV3")
                .Columns("LAST_CONV").Expression = "PARENT(BETCONV0_BETCONV3).LAST_CONV"
                .Columns("SECS").Expression = "PARENT(BETCONV0_BETCONV3).SECS"
                .Columns("ROW_COUNT").Expression = "PARENT(BETCONV0_BETCONV3).ROW_COUNT"
                .Columns("SQL_ERROR").Expression = "PARENT(BETCONV0_BETCONV3).SQL_ERROR"
            End With

            With .Tables("BETCONVX")
                .Columns("LAST_CONV").Expression = "MAX(CHILD(BETCONVX_BETCONV3).LAST_CONV)"
                .Columns("SECS").Expression = "SUM(CHILD(BETCONVX_BETCONV3).SECS)"
                .Columns("ROW_COUNT").Expression = "SUM(CHILD(BETCONVX_BETCONV3).ROW_COUNT)"
                .Columns("SQL_ERROR").Expression = "MAX(CHILD(BETCONVX_BETCONV3).SQL_ERROR)"
            End With

            ASCMAIN1.sql = "Select PROGRAM_CODE from MGTPROG1"
            Create_TDA(.Tables.Add, "MGTPROG1", "**", 0, False, "", 1)
            Fill_Records("MGTPROG1")

            For Each TABLE_NAME As String In New String() {"BETCONVS", "BETCONVC"}
                With .Tables.Add(TABLE_NAME)
                    .Columns.Add("SCHEMA_NAME")
                    .Columns.Add("TABLE_NAME")
                    .Columns.Add("COLUMN_NAME")
                    .Columns.Add("COLUMN_POS", GetType(System.Int32))
                    .Columns.Add("COLUMN_LENGTH", GetType(System.Int32))
                    .Columns.Add("COLUMN_PRECISION", GetType(System.Int32))
                    .Columns.Add("COLUMN_SCALE", GetType(System.Int32))
                    .Columns.Add("DATA_TYPE")
                    .PrimaryKey = New DataColumn() {.Columns("SCHEMA_NAME"), .Columns("TABLE_NAME"), .Columns("COLUMN_NAME")}
                End With
            Next

            With .Tables.Add("BETTABLX")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("USED", GetType(System.Int32))
                .Columns.Add("COLS", GetType(System.Int32))
                .Columns.Add("ISSUES", GetType(System.Int32))
                .PrimaryKey = New DataColumn() {.Columns("TABLE_NAME")}
            End With

            With .Tables.Add("BETCONVD")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("USED", GetType(System.Int32))
                .Columns.Add("ISSUES", GetType(System.Int32))
                .Columns.Add("DEFINITION")
                .PrimaryKey = New DataColumn() {.Columns("TABLE_NAME"), .Columns("COLUMN_NAME")}
            End With

            With .Tables.Add("BETCONVT")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("ROW_NO")
                .PrimaryKey = New DataColumn() {.Columns("TABLE_NAME"), .Columns("COLUMN_NAME"), .Columns("ROW_NO")}
            End With

            Create_Relation("BETCONVD", "BETCONVT", "TABLE_NAME,COLUMN_NAME")
        End With

        grdBETCONV3.DataSource = dst.Tables("BETCONV3")
        grdBETCONVS.DataSource = dst.Tables("BETCONVS")

        grdBETTABLX.DataSource = dst.Tables("BETTABLX")
        grdBETCONVD.DataSource = dst.Tables("BETCONVD")
        grdBETCONVX.DataSource = dst.Tables("BETCONVX")

        grdBETCONVC.DataSource = dst.Tables("BETCONVC")

        grdBETCONV0.DataSource = dst.Tables("BETCONV0")

        Create_Summary(grdBETCONV3, "MYS_SCHEMA_NAME", "Count")

        Create_Summary(grdBETCONV0, "TABLE_NAME", "Count")
        Create_Summary(grdBETCONV0, "ROW_COUNT")
        Create_Summary(grdBETCONV0, "SECS")

        Create_Summary(grdBETTABLX, "TABLE_NAME", "Count")

        Create_Summary(grdBETCONVX, "ORA_TABLE_NAME", "Count")
        Create_Summary(grdBETCONVX, "ROW_COUNT")
        Create_Summary(grdBETCONVX, "SECS")

        grdBETTABLX.DisplayLayout.UseFixedHeaders = True
        With grdBETTABLX.DisplayLayout.Bands(0)
            .Columns("TABLE_NAME").Header.Fixed = True
            .Columns("USED").Header.Fixed = True
            .Columns("COLS").Header.Fixed = True

            .Columns("TABLE_NAME").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("USED").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("COLS").CellAppearance.BackColor = Drawing.Color.LightBlue
        End With

        grdBETCONVD.DisplayLayout.UseFixedHeaders = True
        With grdBETCONVD.DisplayLayout.Bands(0)
            .Columns("COLUMN_NAME").Header.Fixed = True
            .Columns("USED").Header.Fixed = True
            .Columns("DEFINITION").Header.Fixed = True

            .Columns("COLUMN_NAME").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("USED").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("DEFINITION").CellAppearance.BackColor = Drawing.Color.LightBlue
        End With

        grdBETCONV3.DisplayLayout.UseFixedHeaders = True
        With grdBETCONV3.DisplayLayout.Bands(0)
            .Columns("MYS_SCHEMA_NAME").Header.Fixed = True
            .Columns("MYS_TABLE_NAME").Header.Fixed = True
            .Columns("MYS_SCHEMA_NAME").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("MYS_TABLE_NAME").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("MYS_SCHEMA_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("MYS_TABLE_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        Fill_Records("BETCONV0")
        Sort_grdColumns(grdBETCONV0, "SCHEMA_NAME,TABLE_NAME")

        'Load_Dataset()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("ORA_TABLE_NAME", True)
                ASCMAIN1.sql = "Select * from USER_TABLES where TABLE_NAME = :PARM1"
                Dim ORA_TABLE_NAME As String = Absx1.txtFor("ORA_TABLE_NAME").Text
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ORA_TABLE_NAME})
                If row Is Nothing Then
                    EMsg &= vbCr & "Cannot find table " & ORA_TABLE_NAME & " in Database"
                End If

            Case "Edit"
                Validate_Code("ORA_TABLE_NAME")

            Case "Load"
                Validate_Code("ORA_TABLE_NAME")

            Case "Update"
                If Absx1.txtFor("ORA_TABLE_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Enter a Description for this Table"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Generate SQL"
                Generate_SQL()

            Case "Convert to Oracle"
                If Not ScreenMode Then
                    Dim TABLE_NAMEs As New List(Of String)
                    For Each row As DataRow In dst.Tables("BETCONVX").Select("", "ORA_TABLE_NAME")
                        TABLE_NAMEs.Add(row.Item(0))
                    Next
                    For Each TABLE_NAME As String In TABLE_NAMEs
                        Absx1.txtFor("ORA_TABLE_NAME").Text = TABLE_NAME
                        Click_Command("Load")
                        Application.DoEvents()
                        Convert_to_Oracle()
                        Click_Command("Done")
                    Next

                    MsgBox("Conversion to Oracle is Complete - Please check SQL ERROR column for any issues", MsgBoxStyle.OkOnly, "Success")
                Else
                    Convert_to_Oracle()
                End If

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "L" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    .Items("Load").Visible = (EntryMode = "L" Or Not ScreenMode)
                    .Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode

                    .Items("New").Visible = (EntryMode <> "L" Or Not ScreenMode)
                    '.Items("Edit").Visible = (EntryMode <> "L" Or Not ScreenMode)

                    .Items("Update").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Delete").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Cancel").Visible = (Not (EntryMode = "L") Or Not ScreenMode)

                    .Items("Generate SQL").Visible = ScreenMode ' (EntryMode = "L" And ScreenMode)
                    .Items("Convert to Oracle").Visible = True ' ScreenMode ' (EntryMode = "L" And ScreenMode)
                End With
                .Groups("Conversion List").Visible = False
                '.Groups("Generate Options").Visible = (EntryMode = "L" And ScreenMode)
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        If ScreenMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("ORA_TABLE_DESC"), Not ScreenMode)
        End If

        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_ERROR").Hidden = True
        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_STMT").Hidden = True
        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_TEST").Hidden = True

        SplitContainer1.Visible = tf
        'grdBETCONVX.Visible = Not tf
        tabTables.Visible = Not tf

        With grdBETCONV3.DisplayLayout.Override
            If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.No ' FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        If ScreenMode Then
            grdBETCONV0.Parent = grpBETCONV0_B
            chkShowAll.Checked = False

            Show_Filter(grdBETCONV0, True)
        Else
            Clear_Record()
            grdBETCONV0.Parent = SplitContainer4.Panel1
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"BETCONV1", "BETCONV2", "BETCONVS", "BETCONV3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdBETDATAS.DataSource = Nothing
        grdBETDATAD.DataSource = Nothing

        Absx1.txtFor("ORA_TABLE_NAME").Text = ""
        Setup_tabTables()
        Setup_BETCONV0()

        Fill_Records("BETCONVX")
        Fill_Records("BETCONV3", , , "Select * from BETCONV3")
        Sort_grdColumns(grdBETCONVX, "ORA_TABLE_NAME")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        dst.Tables("BETCONVS").Rows.Clear()
        dst.Tables("BETCONVC").Rows.Clear()

        ALL_SCHEMAs.Clear()

        If EntryMode = "N" Then
            rowBETCONV1 = dst.Tables("BETCONV1").NewRow
            rowBETCONV1.Item("ORA_TABLE_NAME") = HFs("ORA_TABLE_NAME")
            rowBETCONV1.Item("ORA_TABLE_DESC") = HFs("ORA_TABLE_DESC")
            dst.Tables("BETCONV1").Rows.Add(rowBETCONV1)

            Dim rowBETCONVX As DataRow = dst.Tables("BETCONVX").NewRow
            rowBETCONVX.Item("ORA_TABLE_NAME") = HFs("ORA_TABLE_NAME")
            rowBETCONVX.Item("ORA_TABLE_DESC") = HFs("ORA_TABLE_DESC")
            dst.Tables("BETCONVX").Rows.Add(rowBETCONVX)

        Else
            rowBETCONV1 = Fill_Record("BETCONV1", HFs("ORA_TABLE_NAME"))
        End If

        ALL_SCHEMAs.Clear()

        For C As Integer = 0 To 199
            grdBETCONV3.DisplayLayout.Bands(0).Columns("COL_" & Format(C, "000")).Hidden = True
        Next

        Load_BETCONVS(".", HFs("ORA_TABLE_NAME"))

        Fill_Records("BETCONV2", HFs("ORA_TABLE_NAME"))
        Fill_Records("BETCONV3", HFs("ORA_TABLE_NAME"))

        For Each rowBETCONV3 As DataRow In dst.Tables("BETCONV3").Select("", "MYS_SCHEMA_NAME,MYS_TABLE_NAME")
            Dim MYS_SCHEMA_NAME As String = rowBETCONV3.Item("MYS_SCHEMA_NAME")
            Dim MYS_TABLE_NAME As String = rowBETCONV3.Item("MYS_TABLE_NAME")
            Load_BETCONVS(MYS_SCHEMA_NAME, MYS_TABLE_NAME)
        Next

        Dim ST As String = ""
        For Each rowBETCONV2 As DataRow In dst.Tables("BETCONV2").Select("", "MYS_SCHEMA_NAME,MYS_TABLE_NAME")

            Dim MYS_SCHEMA_NAME As String = rowBETCONV2.Item("MYS_SCHEMA_NAME")
            Dim MYS_TABLE_NAME As String = rowBETCONV2.Item("MYS_TABLE_NAME")
            Dim ORA_COLUMN_NAME As String = rowBETCONV2.Item("ORA_COLUMN_NAME")
            Dim Cx As Integer = ORA_TBL_SCHEMA(ORA_COLUMN_NAME)
            Dim rowBETCONV3 As DataRow = Nothing
            If ST <> MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME Then
                rowBETCONV3 = dst.Tables("BETCONV3").Rows.Find(New String() {HFs("ORA_TABLE_NAME"), MYS_SCHEMA_NAME, MYS_TABLE_NAME})
            End If
            rowBETCONV3.Item("COL_" & Format(Cx, "000")) = rowBETCONV2.Item("SQL_EXPRESSION")
        Next

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        dst.Tables("BETCONV2").Rows.Clear()
        For Each rowBETCONV3 As DataRow In dst.Tables("BETCONV3").Select()

            Dim MYS_SCHEMA_NAME As String = rowBETCONV3.Item("MYS_SCHEMA_NAME")
            Dim MYS_TABLE_NAME As String = rowBETCONV3.Item("MYS_TABLE_NAME")
            For Each COLUMN_NAME As String In ORA_TBL_SCHEMA.Keys
                If COLUMN_NAME = "PROGRAM_CODE" Then
                Else
                    Dim SQL_EXPRESSION As String = rowBETCONV3.Item("COL_" & Format(ORA_TBL_SCHEMA(COLUMN_NAME), "000")) & ""
                    If SQL_EXPRESSION <> "" Then
                        Dim rowBETCONV2 As DataRow = dst.Tables("BETCONV2").NewRow
                        rowBETCONV2.Item("ORA_TABLE_NAME") = HFs("ORA_TABLE_NAME")
                        rowBETCONV2.Item("ORA_COLUMN_NAME") = COLUMN_NAME
                        rowBETCONV2.Item("MYS_SCHEMA_NAME") = MYS_SCHEMA_NAME
                        rowBETCONV2.Item("MYS_TABLE_NAME") = MYS_TABLE_NAME
                        rowBETCONV2.Item("SQL_EXPRESSION") = SQL_EXPRESSION
                        dst.Tables("BETCONV2").Rows.Add(rowBETCONV2)
                    End If
                End If
            Next
        Next

        Update_Record_TDA("BETCONV1")
        Update_Record_TDA("BETCONV2", "ORA_TABLE_NAME = '" & HFs("ORA_TABLE_NAME") & "'")
        Update_Record_TDA("BETCONV3")
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdBETCONV3, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdBETCONVX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdBETCONV0, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Refresh List from mySQL", "Select Selected", "Select All")
        Load_Popup_Menu(grdBETTABLX, "SSBBB", "Show Filter", "Show Pins", "Refresh Matrix from List", "Get Data", "Generate Oracle DDL")
        Load_Popup_Menu(grdBETCONVD, "SSS", "Show Filter", "Show Pins", "Show Issues Only")
        Load_Popup_Menu(grdBETDATAD, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Data")
        Load_Popup_Menu(grdBETDATAS, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Data")
    End Sub


    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


                Case "grdBETCONV0"
                    DirectCast(tlb_pop.Tools("Refresh List from mySQL"), UltraWinToolbars.ButtonTool).SharedProps.Visible = Not ScreenMode
                    DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool).SharedProps.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                    DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool).SharedProps.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Refresh List from mySQL"
                Refresh_List_from_mySQL()

            Case "Refresh Matrix from List"
                Refresh_Matrix_from_List()

            Case "Show Issues Only"
                Setup_BETCONVD()
                grdBETCONVD.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)

            Case "Select Selected"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Adding the Selected Tables")
                For Each grow As UltraWinGrid.UltraGridRow In grdBETCONV0.Selected.Rows
                    Select_Table(grow.Cells("SCHEMA_NAME").Value, grow.Cells("TABLE_NAME").Value)
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Select All"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Adding All Tables")
                For Each grow As UltraWinGrid.UltraGridRow In grdBETCONV0.Rows
                    If grow.IsFilteredOut Then
                    Else
                        Select_Table(grow.Cells("SCHEMA_NAME").Value, grow.Cells("TABLE_NAME").Value)
                    End If
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Show Data"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Data")
                If grd.Name = "grdBETDATAD" Then
                    ASCMAIN1.sql = "Select * from " & HFs("ORA_TABLE_NAME")
                    grdBETDATAD.DataSource = Nothing
                    grdBETDATAD.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                    grdBETDATAD.DataSource = ASCDATA1.GetDataTable
                    grdBETDATAD.Text = HFs("ORA_TABLE_NAME") & "; " & Format(Val(grdBETDATAD.Rows.Count), "#,##0") & " Rows"
                ElseIf grd.Name = "grdBETDATAS" Then
                    If grdBETCONV3.ActiveRow IsNot Nothing Then
                        Dim SCHEMA_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_SCHEMA_NAME").Value
                        Dim TABLE_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_TABLE_NAME").Value
                        ASCMAIN1.sql = "Select * from " & SCHEMA_NAME & "." & TABLE_NAME
                        grdBETDATAS.DataSource = Nothing
                        grdBETDATAS.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                        grdBETDATAS.DataSource = Fill_Table_from_MYSQL(Nothing, ASCMAIN1.sql)
                        grdBETDATAS.Text = SCHEMA_NAME & "." & TABLE_NAME & "; " & Format(Val(grdBETDATAS.Rows.Count), "#,##0") & " Rows"
                    End If
                End If
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
            '    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")
            Case "Get Data"
                Get_Data(grd.ActiveRow.Cells("TABLE_NAME").Value)

            Case "Generate Oracle DDL"
                Generate_Oracle_DDL(grd.ActiveRow.Cells("TABLE_NAME").Value)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If ScreenMode Then
            Exit Sub
        End If

        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ORA_TABLE_NAME"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ORA_TABLE_NAME"
                Click_Command("Load")
        End Select
    End Sub

#End Region

    Private Sub grdBETCONVX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdBETCONVX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ORA_TABLE_NAME").Text = e.Row.Cells("ORA_TABLE_NAME").Value
            Click_Command("Load")
        End If
    End Sub

    Function Fill_Table_from_MYSQL(ByVal TBL As DataTable, ByVal SQL As String) As DataTable
        If TBL Is Nothing Then TBL = New DataTable
        Using CON As New MySql.Data.MySqlClient.MySqlConnection(CONSTR)
            Dim ADA As New MySql.Data.MySqlClient.MySqlDataAdapter(SQL, CON)
            ADA.Fill(TBL)
        End Using
        Return TBL
    End Function

    Sub Setup_grdBETCONVS()
        If grdBETCONV3.ActiveRow Is Nothing OrElse Not grdBETCONV3.ActiveRow.IsDataRow Then
            grdBETDATAS.Visible = False
        Else
            grdBETDATAS.Visible = True
            Dim MYS_SCHEMA_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_SCHEMA_NAME").Value
            Dim MYS_TABLE_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_TABLE_NAME").Value
            tblBETCONVS.Rows.Clear()
            Fill_Table_from_MYSQL(tblBETCONVS, Replace(Replace(sqlBETCONVS, "SCHEMA_NAME", MYS_SCHEMA_NAME), "TABLE_NAME", MYS_TABLE_NAME))
        End If
    End Sub

    Sub Generate_SQL()

        Using CON As New MySql.Data.MySqlClient.MySqlConnection(CONSTR)
            Dim SQLt As String = ""
            For Each rowBETCONV3 As DataRow In dst.Tables("BETCONV3").Select("")
                Dim SQL As String = ""
                Dim MYS_SCHEMA_NAME As String = rowBETCONV3.Item("MYS_SCHEMA_NAME")
                Dim MYS_TABLE_NAME As String = rowBETCONV3.Item("MYS_TABLE_NAME")
                For Each COLUMN_NAME As String In ORA_TBL_SCHEMA.Keys
                    Dim CX As Integer = ORA_TBL_SCHEMA(COLUMN_NAME)

                    If COLUMN_NAME = "PROGRAM_CODE" Then
                    Else
                        Dim SQL_EXPRESSION As String = rowBETCONV3.Item("COL_" & Format(CX, "000")) & ""
                        If SQL_EXPRESSION = "" Then
                            SQL_EXPRESSION = "NULL"
                        End If
                        If SQL_EXPRESSION = COLUMN_NAME Then
                            SQL &= "," & SQL_EXPRESSION
                        Else
                            SQL &= "," & SQL_EXPRESSION & " " & IIf("COLUMN_NAME" = "access", "LOGIN_access", Mid(COLUMN_NAME, 1, 30))
                        End If
                    End If
                Next
                Dim SQL_STMT As String = "Select '" & MYS_SCHEMA_NAME.ToUpper & "' PROGRAM_CODE" & SQL & " from " & MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME
                rowBETCONV3.Item("SQL_STMT") = SQL_STMT

                rowBETCONV3.Item("SQL_TEST") = ""

                Try
                    Dim TBL As New DataTable
                    Using ADA As New MySql.Data.MySqlClient.MySqlDataAdapter(SQL_STMT & " LIMIT 1", CON)
                        ADA.Fill(TBL)
                    End Using
                Catch ex As Exception
                    rowBETCONV3.Item("SQL_TEST") = ex.Message
                End Try

                SQL = "Insert into " & HFs("ORA_TABLE_NAME") & " " & SQL_STMT
                SQLt &= SQL & vbCrLf
            Next
        End Using
        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_STMT").Hidden = False
        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_TEST").Hidden = False
        Update_Record_TDA("BETCONV3")
    End Sub

    Sub Convert_to_Oracle()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data for " & HFs("ORA_TABLE_NAME"))

        Generate_SQL()

        If Not TDAs.ContainsKey(HFs("ORA_TABLE_NAME")) Then
            Create_TDA(dst.Tables.Add, HFs("ORA_TABLE_NAME"), "*")
        End If

        Create_BAs(HFs("ORA_TABLE_NAME"))
        ASCDATA1.ExecuteSQL("Truncate Table " & HFs("ORA_TABLE_NAME"))

        For Each rowBETCONV3 As DataRow In dst.Tables("BETCONV3").Select()
            Dim MYS_SCHEMA_NAME As String = rowBETCONV3.Item("MYS_SCHEMA_NAME")
            Dim MYS_TABLE_NAME As String = rowBETCONV3.Item("MYS_TABLE_NAME")
            Dim SQL_STMT As String = rowBETCONV3.Item("SQL_STMT") & ""
            If SQL_STMT <> "" Then
                Try
                    Load_Oracle_Table(HFs("ORA_TABLE_NAME"), SQL_STMT, MYS_SCHEMA_NAME, MYS_TABLE_NAME)
                Catch ex As Exception
                    Dim rowBETCONV0 As DataRow = dst.Tables("BETCONV0").Rows.Find(New String() {MYS_SCHEMA_NAME, MYS_TABLE_NAME})
                    rowBETCONV0.Item("SQL_ERROR") = Mid(ex.Message, 1, 4000)
                    If Not rowBETCONV0.Item("CONVERT") & "" = "1" Then Continue For
                    Update_Record_TDA("BETCONV0")
                End Try
            End If
        Next

        grdBETCONV3.DisplayLayout.Bands(0).Columns("SQL_ERROR").Hidden = False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Oracle_Table(ORA_TABLE_NAME As String, SQL_STMT As String, MYS_SCHEMA_NAME As String, MYS_TABLE_NAME As String)

        ASCMAIN1.Progress("Now Converting " & MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME & " -> " & ORA_TABLE_NAME)

        Dim TBL As DataTable = Fill_Table_from_MYSQL(Nothing, SQL_STMT)
        Dim tblO As DataTable = dst.Tables(ORA_TABLE_NAME)

        If MYS_SCHEMA_NAME = "zometa" And MYS_TABLE_NAME = "eob" Then
            Dim ID As Integer = 0
            For Each row As DataRow In TBL.Rows
                ID += 1
                row.Item("ID") = ID
            Next
        End If

        Dim rowBETCONV0 As DataRow = dst.Tables("BETCONV0").Rows.Find(New String() {MYS_SCHEMA_NAME, MYS_TABLE_NAME})

        Dim START_TIME As Date = Now
        Dim dtcols As New List(Of String)
        Dim tscols As New List(Of String)
        For Each dc As DataColumn In TBL.Columns
            Select Case dc.DataType.ToString
                Case "MySql.Data.Types.MySqlDateTime"
                    dtcols.Add(dc.ColumnName)
                Case "System.TimeSpan"
                    tscols.Add(dc.ColumnName)
            End Select
        Next

        Dim R As Int64 = 0
        Dim tsd As New Dictionary(Of String, Date)

        For Each rowm As DataRow In TBL.Rows
            Dim rowO As DataRow = tblO.NewRow
            If dtcols.Count <> 0 Then
                For Each dtcol As String In dtcols
                    ' If rowm.Item("ID") = 86568 Then Stop
                    Dim dtv As String = rowm.Item(dtcol).ToString
                    If dtv = "" OrElse Not IsDate(dtv) Then
                        rowm.Item(dtcol) = DBNull.Value
                    End If
                Next
            End If
            If tscols.Count <> 0 Then
                tsd = New Dictionary(Of String, Date)
                For Each tscol As String In tscols
                    ' If rowm.Item("ID") = 86568 Then Stop
                    Dim tsv As String = rowm.Item(tscol).ToString
                    Dim ts As TimeSpan = rowm.Item(tscol)
                    Dim dt As Date
                    If rowm.Item("VISITDATE") & "" = "" Then
                        dt = Now.Date
                    Else
                        dt = rowm.Item("VISITDATE")
                    End If
                    dt = dt.AddTicks(ts.Ticks) ' dt.AddHours(ts.Hours).AddMinutes(ts.Minutes).AddSeconds(ts.Seconds)
                    tsd.Add(tscol, dt)

                    rowm.Item(tscol) = DBNull.Value
                    'If tsv = "" OrElse Not IsDate(tsv) Then
                    '    rowm.Item(tscol) = DBNull.Value
                    'End If
                Next
            End If

            rowO.ItemArray = rowm.ItemArray
            If tscols.Count <> 0 Then
                For Each tscol As String In tscols
                    rowO.Item(tscol) = tsd(tscol)
                Next
            End If

            If MYS_TABLE_NAME = "cardinfo" Then
                Dim s As String = ""
                If rowm.Item("MC") IsNot Null Then
                    Dim B() As Byte = rowm.Item("MC")
                    Dim enc As New System.Text.UTF8Encoding()
                    s = enc.GetString(B)
                End If

                rowO.Item("MC") = s
            End If

            R += 1
            tblO.Rows.Add(rowO)
            If R Mod 1000 = 0 Then
                Update_BAs(HFs("ORA_TABLE_NAME"))
                ASCMAIN1.Progress("-", CStr(R))
                tblO.Rows.Clear()
            End If
        Next
        If R Mod 1000 <> 0 Then
            Update_BAs(HFs("ORA_TABLE_NAME"))
            tblO.Rows.Clear()
        End If

        ASCMAIN1.sql = "Select * from " & MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME
        Dim tblSchema As DataTable
        Dim KEY_COUNT As Integer = 0
        Dim COL_COUNT As Integer = 0
        Using CON As New MySql.Data.MySqlClient.MySqlConnection(CONSTR)
            CON.Open()
            Using ADA As New MySql.Data.MySqlClient.MySqlDataAdapter(ASCMAIN1.sql, CON)
                Using RDR As MySql.Data.MySqlClient.MySqlDataReader = ADA.SelectCommand.ExecuteReader
                    tblSchema = RDR.GetSchemaTable()
                    KEY_COUNT = tblSchema.Select("IsKey").Length
                    COL_COUNT = tblSchema.Columns.Count
                End Using
            End Using
            CON.Close()
        End Using

        Dim SECS As Integer = Now.Subtract(START_TIME).TotalSeconds
        rowBETCONV0.Item("ROW_COUNT") = TBL.Rows.Count
        rowBETCONV0.Item("COL_COUNT") = COL_COUNT
        rowBETCONV0.Item("KEY_COUNT") = KEY_COUNT
        rowBETCONV0.Item("LAST_CONV") = DATETIME_STAMP
        rowBETCONV0.Item("SECS") = SECS
        rowBETCONV0.Item("SQL_ERROR") = DBNull.Value
        Update_Record_TDA("BETCONV0")
    End Sub

    Sub Load_Dataset()
        ASCMAIN1.sql = "Select TABLE_NAME from USER_TABLES where TABLE_NAME like 'BET%' and TABLE_NAME not like 'BETCONV%'"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Rows
            Dim TABLE_NAME As String = ROW.Item(0)
            ASCMAIN1.sql = "Select * from " & TABLE_NAME
             Create_TDA(dst.Tables.Add, TABLE_NAME, "**", 0, False)
        Next
    End Sub

    Sub Refresh_List_from_mySQL()

        Dim sql As String = ""
        Using CON As New MySql.Data.MySqlClient.MySqlConnection(CONSTR)
            CON.Open()

            sql = "Select Distinct TABLE_SCHEMA " _
            & " from INFORMATION_SCHEMA.TABLES where TABLE_SCHEMA " _
            & " NOT in ('sakila','test','world','information_schema','mysql','performance','performance_schema')"

            Dim tblSCHEMAS As New DataTable
            Using ADA As New MySql.Data.MySqlClient.MySqlDataAdapter(sql, CON)
                ADA.Fill(tblSCHEMAS)
            End Using

            For Each rowSCHEMA As DataRow In tblSCHEMAS.Rows
                Dim TABLE_SCHEMA As String = rowSCHEMA.Item(0)

                sql = "Select TABLE_NAME, TABLE_ROWS from INFORMATION_SCHEMA.TABLES" _
                & " where TABLE_SCHEMA = '" & TABLE_SCHEMA & "'"

                Dim tblTABLES As New DataTable
                Using ADA2 As New MySql.Data.MySqlClient.MySqlDataAdapter(sql, CON)
                    ADA2.Fill(tblTABLES)
                End Using

                For Each rowTABLE As DataRow In tblTABLES.Rows
                    Dim TABLE_NAME As String = rowTABLE.Item(0)
                    'If TABLE_NAME = "treatment" Then Stop
                    Dim rowBETCONV0 As DataRow = dst.Tables("BETCONV0").Rows.Find(New String() {TABLE_SCHEMA, TABLE_NAME})
                    If rowBETCONV0 Is Nothing Then

                        Dim TBL As New DataTable
                        sql = "Select * from " & TABLE_SCHEMA & "." & TABLE_NAME
                        Using ADA3 As New MySql.Data.MySqlClient.MySqlDataAdapter(sql, CON)
                            Using RDR As MySql.Data.MySqlClient.MySqlDataReader = ADA3.SelectCommand.ExecuteReader
                                TBL = RDR.GetSchemaTable()
                            End Using
                        End Using
                        TBL.PrimaryKey = New DataColumn() {TBL.Columns("ColumnName")}

                        rowBETCONV0 = dst.Tables("BETCONV0").NewRow
                        With rowBETCONV0
                            .Item("SCHEMA_NAME") = TABLE_SCHEMA
                            .Item("TABLE_NAME") = TABLE_NAME
                            .Item("ROW_COUNT") = rowTABLE.Item("TABLE_ROWS")
                            .Item("KEY_COUNT") = TBL.Select("IsKey = True").Length
                            .Item("COL_COUNT") = TBL.Rows.Count
                        End With
                        dst.Tables("BETCONV0").Rows.Add(rowBETCONV0)
                    End If
                Next
            Next
            CON.Close()
        End Using

        Update_Record_TDA("BETCONV0")
        Sort_grdColumns(grdBETCONV0, "SCHEMA_NAME,TABLE_NAME")
        MsgBox("List of Tables has been Refreshed", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Refresh_Matrix_from_List()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Tables")

        grdBETTABLX.DataSource = Nothing

        SCHEMAs = New Dictionary(Of String, Integer)
        Dim tbl As New DataTable
        tbl.Columns.Add("TABLE_NAME")
        tbl.Columns.Add("COLS", GetType(System.Int32))
        tbl.Columns.Add("USES", GetType(System.Int32))
        tbl.PrimaryKey = New DataColumn() {tbl.Columns("TABLE_NAME"), tbl.Columns("COLS")}

        Dim COLs_original As Integer = 4
        grdBETTABLX.DisplayLayout.Bands(0).Summaries.Clear()

        With dst.Tables("BETTABLX")
            .Rows.Clear()

            If .Columns.Count > COLs_original Then
                For C As Integer = .Columns.Count - 1 To COLs_original Step -1
                    .Columns.Remove(.Columns(C).ColumnName)
                Next
            End If

            For Each row As DataRow In ASCDATA1.SelectDistinct _
                    (dst.Tables("BETCONV0").Select("CONVERT = '1'", "SCHEMA_NAME"), New String() {"SCHEMA_NAME"}).Select
                Dim SCHEMA_NAME As String = row.Item("SCHEMA_NAME")
                .Columns.Add(SCHEMA_NAME)
                SCHEMAs.Add(SCHEMA_NAME, 0)
            Next

            For Each rowBETCONV0 As DataRow In dst.Tables("BETCONV0").Select("CONVERT = '1'", "SCHEMA_NAME,TABLE_NAME")
                Dim SCHEMA_NAME As String = rowBETCONV0.Item("SCHEMA_NAME")
                Dim TABLE_NAME As String = rowBETCONV0.Item("TABLE_NAME")
                Dim ROWS As Integer = Val(rowBETCONV0.Item("ROW_COUNT") & "")
                Dim COLS As Integer = Val(rowBETCONV0.Item("COL_COUNT") & "")
                Dim RC As String = CStr(ROWS) & "/" & CStr(COLS)

                ASCMAIN1.Progress("-", SCHEMA_NAME & "." & TABLE_NAME)

                Dim row As DataRow = tbl.Rows.Find(New Object() {TABLE_NAME, COLS})
                If row Is Nothing Then
                    tbl.Rows.Add(New Object() {TABLE_NAME, COLS, 1})
                Else
                    row.Item("USES") = Val(row.Item("USES") & "") + 1
                End If

                Dim rowBETTABLX As DataRow = .Rows.Find(TABLE_NAME)
                If rowBETTABLX Is Nothing Then
                    rowBETTABLX = .NewRow
                    rowBETTABLX.Item("TABLE_NAME") = TABLE_NAME
                    .Rows.Add(rowBETTABLX)
                End If
                Dim USED As Integer = Val(rowBETTABLX.Item("USED") & "")
                rowBETTABLX.Item(SCHEMA_NAME) = RC
                rowBETTABLX.Item("USED") = USED + 1
            Next

            For Each rowBETTABLX As DataRow In .Select()
                Dim TABLE_NAME As String = rowBETTABLX.Item("TABLE_NAME")
                Dim rows() As DataRow = tbl.Select("TABLE_NAME = '" & TABLE_NAME & "'", "USES DESC")
                If rows.Length <> 0 Then
                    Dim COLS As Integer = rows(0).Item("COLS")
                    rowBETTABLX.Item("COLS") = COLS

                    Dim ISSUES As Int32 = 0
                    For Each SCHEMA_NAME As String In SCHEMAs.Keys
                        Dim RC As String = rowBETTABLX.Item(SCHEMA_NAME) & ""
                        If RC <> "" Then
                            Dim COLSx As Integer = Val(Split(RC & "/", "/")(1))
                            If COLSx <> COLS Then
                                ISSUES += 1
                            End If
                        End If
                    Next
                    If ISSUES <> 0 Then rowBETTABLX.Item("ISSUES") = ISSUES
                End If
            Next
        End With

        grdBETTABLX.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdBETTABLX.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show

        grdBETTABLX.DataSource = dst.Tables("BETTABLX")
        Create_Summary(grdBETTABLX, "TABLE_NAME", "Count")
        For Each SCHEMA As String In SCHEMAs.Keys
            Create_Summary(grdBETTABLX, SCHEMA, "Custom")
        Next

        With grdBETTABLX.DisplayLayout.Bands(0)
            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                With .Columns(SCHEMA_NAME)
                    .CellAppearance.TextHAlign = HAlign.Center
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .Width = 80
                End With
            Next
        End With

        grdBETTABLX.DisplayLayout.UseFixedHeaders = True
        With grdBETTABLX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"TABLE_NAME", "USED", "COLS", "ISSUES"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Caption = ASCMAIN1.Make_Caption(Split(COLUMN_NAME & "_", "_")(0))
            Next
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Definitions")

        grdBETCONVD.DataSource = Nothing

        Dim tblD As New DataTable
        tblD.Columns.Add("TABLE_NAME")
        tblD.Columns.Add("COLUMN_NAME")
        tblD.Columns.Add("DEFINITION")
        tblD.Columns.Add("USES", GetType(System.Int32))
        tblD.PrimaryKey = New DataColumn() {tblD.Columns("TABLE_NAME"), tblD.Columns("COLUMN_NAME"), tblD.Columns("DEFINITION")}

        Dim COLs_original_BETCONVT As Integer = 3
        With dst.Tables("BETCONVT")
            .Rows.Clear()
            If .Columns.Count > COLs_original_BETCONVT Then
                For C As Integer = .Columns.Count - 1 To COLs_original_BETCONVT Step -1
                    .Columns.Remove(.Columns(C).ColumnName)
                Next
            End If
            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                .Columns.Add(SCHEMA_NAME)
            Next
        End With

        Dim COLs_original_BETCONVD As Integer = 5
        grdBETCONVD.DisplayLayout.Bands(0).Summaries.Clear()

        With dst.Tables("BETCONVD")
            .Rows.Clear()

            If .Columns.Count > COLs_original_BETCONVD Then
                For C As Integer = .Columns.Count - 1 To COLs_original_BETCONVD Step -1
                    .Columns.Remove(.Columns(C).ColumnName)
                Next
            End If

            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                .Columns.Add(SCHEMA_NAME)
            Next

            dst.Tables("BETCONVS").Rows.Clear()
            ALL_SCHEMAs.Clear()

            For Each rowBETCONV0 As DataRow In dst.Tables("BETCONV0").Select("CONVERT = '1'", "SCHEMA_NAME")
                Dim SCHEMA_NAME As String = rowBETCONV0.Item("SCHEMA_NAME")
                Dim TABLE_NAME As String = rowBETCONV0.Item("TABLE_NAME")

                ASCMAIN1.Progress("-", SCHEMA_NAME & "." & TABLE_NAME)

                Load_BETCONVS(SCHEMA_NAME, TABLE_NAME)

                For Each rowBETCONVS As DataRow In dst.Tables("BETCONVS").Select _
                        (String.Format("SCHEMA_NAME = '{0}' and TABLE_NAME = '{1}'", SCHEMA_NAME, TABLE_NAME))

                    Dim COLUMN_NAME As String = rowBETCONVS.Item("COLUMN_NAME")
                    Dim DATA_TYPE As String = rowBETCONVS.Item("DATA_TYPE")
                    Dim SPS As String = "(" & rowBETCONVS.Item("COLUMN_LENGTH") & ")"
                    If rowBETCONVS.Item("COLUMN_SCALE") & "" <> "" And DATA_TYPE <> "System.String" Then
                        SPS = "(" & rowBETCONVS.Item("COLUMN_PRECISION") & "," & rowBETCONVS.Item("COLUMN_SCALE") & ")"
                    End If
                    If SPS = "(0,0)" Then SPS = ""
                    Dim DEFINITION As String = DATA_TYPE
                    If DATA_TYPE.EndsWith("Date") Or DATA_TYPE.EndsWith("DateTime") Then
                    Else
                        DEFINITION = DEFINITION & SPS
                    End If

                    Dim row As DataRow = tblD.Rows.Find(New Object() {TABLE_NAME, COLUMN_NAME, DEFINITION})
                    If row Is Nothing Then
                        tblD.Rows.Add(New Object() {TABLE_NAME, COLUMN_NAME, DEFINITION, 1})
                    Else
                        row.Item("USES") = Val(row.Item("USES") & "") + 1
                    End If

                    Dim rowBETCONVD As DataRow = .Rows.Find(New Object() {TABLE_NAME, COLUMN_NAME})
                    If rowBETCONVD Is Nothing Then
                        rowBETCONVD = .NewRow
                        rowBETCONVD.Item("TABLE_NAME") = TABLE_NAME
                        rowBETCONVD.Item("COLUMN_NAME") = COLUMN_NAME
                        .Rows.Add(rowBETCONVD)
                    End If
                    Dim USED As Integer = Val(rowBETCONVD.Item("USED") & "")
                    rowBETCONVD.Item(SCHEMA_NAME) = DEFINITION
                    rowBETCONVD.Item("USED") = USED + 1
                Next
            Next

            For Each rowBETCONVD As DataRow In .Select()
                Dim TABLE_NAME As String = rowBETCONVD.Item("TABLE_NAME")
                Dim rowBETTABLX As DataRow = dst.Tables("BETTABLX").Rows.Find(TABLE_NAME)
                Dim COLUMN_NAME As String = rowBETCONVD.Item("COLUMN_NAME")
                Dim rows() As DataRow = tblD.Select(String.Format("TABLE_NAME = '{0}' and COLUMN_NAME = '{1}'", TABLE_NAME, COLUMN_NAME), "USES DESC")
                If rows.Length <> 0 Then
                    Dim DEFINITION As String = rows(0).Item("DEFINITION")
                    rowBETCONVD.Item("DEFINITION") = DEFINITION

                    Dim ISSUES As Int32 = 0
                    For Each SCHEMA_NAME As String In SCHEMAs.Keys
                        If rowBETTABLX.Item(SCHEMA_NAME) & "" <> "" Then
                            Dim DEF As String = rowBETCONVD.Item(SCHEMA_NAME) & ""
                            If DEF <> DEFINITION Then
                                ISSUES += 1
                            End If
                        End If
                    Next
                    If ISSUES <> 0 Then rowBETCONVD.Item("ISSUES") = ISSUES
                End If
            Next
        End With

        grdBETCONVD.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdBETCONVD.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show

        grdBETCONVD.DataSource = dst.Tables("BETCONVD")
        Create_Summary(grdBETCONVD, "COLUMN_NAME", "Count")
        For Each SCHEMA As String In SCHEMAs.Keys
            Create_Summary(grdBETCONVD, SCHEMA, "Custom")
        Next

        With grdBETCONVD.DisplayLayout.Bands(0)
            .Columns("DEFINITION").Width = 250
            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                .Columns(SCHEMA_NAME).Width = 250
            Next
        End With

        With grdBETCONVD.DisplayLayout.Bands(1)
            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                .Columns(SCHEMA_NAME).Width = 250
            Next
        End With

        grdBETCONVD.DisplayLayout.UseFixedHeaders = True
        With grdBETCONVD.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"COLUMN_NAME", "USED", "DEFINITION", "ISSUES"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Caption = ASCMAIN1.Make_Caption(Split(COLUMN_NAME & "_", "_")(0))
            Next
            .Columns("TABLE_NAME").Hidden = True
        End With

        grdBETCONVD.DisplayLayout.Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized

        With grdBETCONVD.DisplayLayout.Bands(1)
            For Each COLUMN_NAME As String In New String() {"ROW_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Caption = ASCMAIN1.Make_Caption(Split(COLUMN_NAME & "_", "_")(0))
            Next
            .Columns("TABLE_NAME").Hidden = True
            .Columns("COLUMN_NAME").Hidden = True
            .Columns("ROW_NO").ColSpan = 4

        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Setup_BETCONVD()
    End Sub

    Private Sub _AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdBETTABLX.AfterRowActivate
        If grdBETTABLX.ActiveRow.IsDataRow Then
            Setup_BETCONVD()
        End If
    End Sub

    Sub Setup_BETCONVD()
        If grdBETTABLX.ActiveRow Is Nothing Then
            grdBETCONVD.Visible = False
        Else
            grdBETCONVD.Visible = True

            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                grdBETCONVD.DisplayLayout.Bands(0).Columns(SCHEMA_NAME).Hidden = (grdBETTABLX.ActiveRow.Cells(SCHEMA_NAME).Value & "" = "")
                grdBETCONVD.DisplayLayout.Bands(1).Columns(SCHEMA_NAME).Hidden = (grdBETTABLX.ActiveRow.Cells(SCHEMA_NAME).Value & "" = "")
              Next

            Dim TABLE_NAME As String = grdBETTABLX.ActiveRow.Cells("TABLE_NAME").Value
            Dim dvw As DataView = DirectCast(grdBETCONVD.DataSource, DataTable).DefaultView
            dvw.RowFilter = "TABLE_NAME = '" & TABLE_NAME & "'"
            grdBETCONVD.Text = "mySQL Tables - Definition Comparison - " & TABLE_NAME

        End If
    End Sub

    Private Sub grdBETTABLX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdBETTABLX.InitializeRow

        Dim COLS As Integer = Val(e.Row.Cells("COLS").Value & "")
        For Each SCHEMA_NAME As String In SCHEMAs.Keys
            Dim RC As String = e.Row.Cells(SCHEMA_NAME).Value & ""
            Dim ROWSx As Integer = Val(Split(RC, "/")(0))
            Dim COLSx As Integer = Val(Split(RC & "/", "/")(1))
            If ROWSx <> 0 Then
                e.Row.Cells(SCHEMA_NAME).Appearance.BackColor = Drawing.Color.LightGreen
            End If
            If COLS <> COLSx Then
                e.Row.Cells(SCHEMA_NAME).Appearance.ForeColor = Drawing.Color.Red
            End If
        Next
    End Sub

    Private Sub grdBETCONV0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdBETCONV0.AfterRowActivate
        If Not ScreenMode AndAlso grdBETCONV0.ActiveRow.IsDataRow Then
            lblSchemaName.Text = grdBETCONV0.ActiveRow.Cells("SCHEMA_NAME").Value
            lblTableName.Text = grdBETCONV0.ActiveRow.Cells("TABLE_NAME").Value
            UltraExplorerBar1.Groups("Conversion List").Visible = True

            setup_cmdList_Caption()
        Else
            UltraExplorerBar1.Groups("Conversion List").Visible = False
        End If
    End Sub

    Private Sub grdBETCONV0_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdBETCONV0.DoubleClickRow
        If e.Row.IsDataRow Then
            If ScreenMode Then
                If EntryMode = "E" Or EntryMode = "N" Then
                    Dim SCHEMA_NAME As String = e.Row.Cells("SCHEMA_NAME").Value
                    Dim TABLE_NAME As String = e.Row.Cells("TABLE_NAME").Value

                    Select_Table(SCHEMA_NAME, TABLE_NAME)
                End If
            Else
                Setup_Data_and_Definition(e.Row.Cells("SCHEMA_NAME").Value, e.Row.Cells("TABLE_NAME").Value)
            End If
        End If
    End Sub

    Sub Select_Table(SCHEMA_NAME As String, TABLE_NAME As String)
        Dim rowBETCONV3 As DataRow = dst.Tables("BETCONV3").Rows.Find(New String() {HFs("ORA_TABLE_NAME"), SCHEMA_NAME, TABLE_NAME})
        If rowBETCONV3 Is Nothing Then
            Load_BETCONVS(SCHEMA_NAME, TABLE_NAME)
            rowBETCONV3 = dst.Tables("BETCONV3").Rows.Add(New Object() {HFs("ORA_TABLE_NAME"), SCHEMA_NAME, TABLE_NAME})

            For Each rowBETCONVS As DataRow In dst.Tables("BETCONVS").Select("SCHEMA_NAME = '.'")
                Dim Cx As Integer = Val(rowBETCONVS.Item("COLUMN_POS"))
                Dim COLUMN_NAME As String = rowBETCONVS.Item("COLUMN_NAME")
                If ALL_SCHEMAs(SCHEMA_NAME & "." & TABLE_NAME).ContainsKey(COLUMN_NAME) Then
                    rowBETCONV3.Item("COL_" & Format(Cx, "000")) = rowBETCONVS.Item("COLUMN_NAME")
                End If
            Next
        End If
    End Sub

    Sub Setup_Data_and_Definition(SCHEMA_NAME As String, TABLE_NAME As String)
        If grdBETCONV0.ActiveRow Is Nothing Then 'Not New String() {"TABLE_NAME", "USES", "COLS"}.Contains(e.Cell.Column.Key) Then
            splData_and_Definition.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now setting up data and definition")

            splData_and_Definition.Visible = True
            Dim CMD As String = "SELECT * from " & SCHEMA_NAME & "." & TABLE_NAME

            Dim TBL As New DataTable
            Fill_Table_from_MYSQL(TBL, CMD)

            grdData.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grdData.DataSource = TBL
            grdData.Text = "Data for " & SCHEMA_NAME & "." & TABLE_NAME & " (" & TBL.Rows.Count & " Rows)"

              Dim tblSchema As New DataTable
            With tblSchema
                .Columns.Add("ColumnName")
                .Columns.Add("DataType")
            End With

            For Each dc As DataColumn In TBL.Columns
                tblSchema.Rows.Add(New Object() {dc.ColumnName, dc.DataType.ToString})
            Next

            grdDefinition.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            grdDefinition.DataSource = tblSchema ' COLS
            grdDefinition.Text = "Record Layout for " & SCHEMA_NAME & "." & TABLE_NAME

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Sub Load_BETCONVS(SCHEMA_NAME As String, TABLE_NAME As String)

        Dim TBL_SCHEMA As New DataTable

        Dim SCHEMA As String = SCHEMA_NAME ' Get_Schema(SCHEMA_NAME)
        Dim TBLNAME As String = TABLE_NAME ' Get_Table(TABLE_NAME)

        If SCHEMA_NAME = "." Then
            ASCMAIN1.sql = "Select * from " & TABLE_NAME

            'Dim ada As New Oracle.DataAccess.Client.OracleDataAdapter(ASCMAIN1.sql, ASCMAIN1.oraCon)
            'TBL_SCHEMA = ada.SelectCommand.ExecuteReader.GetSchemaTable
            TBL_SCHEMA = Get_Schema
        Else
            ASCMAIN1.sql = "Select * from " & SCHEMA_NAME & "." & TABLE_NAME
            ASCMAIN1.sql = "Select * from " & SCHEMA & "." & TBLNAME
            Using CON As New MySql.Data.MySqlClient.MySqlConnection(CONSTR)
                CON.Open()
                Try
                    Dim mysada As New MySql.Data.MySqlClient.MySqlDataAdapter(ASCMAIN1.sql, CON)

                    TBL_SCHEMA = mysada.SelectCommand.ExecuteReader.GetSchemaTable
                Catch ex As Exception
                    'HEADS UP
                End Try
                CON.Close()
            End Using
        End If

        Dim TBL_SCHEMAx As New Dictionary(Of String, Integer)

        For Each row As DataRow In TBL_SCHEMA.Rows
            Dim rowBETCONVS As DataRow = dst.Tables("BETCONVS").NewRow
            With rowBETCONVS
                .Item("SCHEMA_NAME") = SCHEMA_NAME
                .Item("TABLE_NAME") = TABLE_NAME
                .Item("COLUMN_NAME") = row.Item("ColumnName")
                .Item("COLUMN_POS") = row.Item("ColumnOrdinal")
                .Item("COLUMN_LENGTH") = row.Item("ColumnSize")
                .Item("COLUMN_PRECISION") = row.Item("NumericPrecision")
                .Item("COLUMN_SCALE") = row.Item("NumericScale")
                .Item("DATA_TYPE") = row.Item("DataType")

                Dim Cx As Integer = Val(rowBETCONVS.Item("COLUMN_POS"))
                TBL_SCHEMAx.Add(.Item("COLUMN_NAME").ToString.ToUpper, Cx)

                If SCHEMA_NAME = "." Then
                    With grdBETCONV3.DisplayLayout.Bands(0).Columns("COL_" & Format(Cx, "000"))
                        .Hidden = (rowBETCONVS.Item("COLUMN_NAME") = "PROGRAM_CODE")
                        .Header.Caption = rowBETCONVS.Item("COLUMN_NAME")
                    End With

                    Dim rowBETCONVC As DataRow = dst.Tables("BETCONVC").NewRow
                    rowBETCONVC.ItemArray = rowBETCONVS.ItemArray
                    dst.Tables("BETCONVC").Rows.Add(rowBETCONVC)
                End If
            End With

            dst.Tables("BETCONVS").Rows.Add(rowBETCONVS)
        Next
        If SCHEMA_NAME = "." Then
            ORA_TBL_SCHEMA = TBL_SCHEMAx
        Else
            ALL_SCHEMAs.Add(SCHEMA_NAME & "." & TABLE_NAME, TBL_SCHEMAx)
        End If
    End Sub

    Private Sub grdBETCONV3_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdBETCONV3.AfterRowActivate
        Setup_BETCONVS()
    End Sub

    Sub Setup_BETCONVS()
        If grdBETCONV3.ActiveRow Is Nothing Then
            grdBETCONVS.Visible = False
        Else
            grdBETCONVS.Visible = True

            Dim MYS_SCHEMA_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_SCHEMA_NAME").Value
            Dim MYS_TABLE_NAME As String = grdBETCONV3.ActiveRow.Cells("MYS_TABLE_NAME").Value
            Dim dvw As DataView = DirectCast(grdBETCONVS.DataSource, DataTable).DefaultView
            dvw.RowFilter = String.Format("SCHEMA_NAME = '{0}' and TABLE_NAME = '{1}'", MYS_SCHEMA_NAME, MYS_TABLE_NAME)
            grdBETCONVS.DisplayLayout.Bands(0).Columns("COLUMN_NAME").Header.Caption = MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME
        End If
    End Sub

    Private Sub grdBETCONV3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdBETCONV3.AfterRowsDeleted
        For Each ST As String In BeforeRowsDeleted
            Dim MYS_SCHEMA_NAME As String = Split(ST, ".")(0)
            Dim MYS_TABLE_NAME As String = Split(ST, ".")(1)
            ASCDATA1.DeleteRows(dst.Tables("BETCONVS"), String.Format("SCHEMA_NAME = '{0}' and TABLE_NAME = '{1}'", MYS_SCHEMA_NAME, MYS_TABLE_NAME))
            ALL_SCHEMAs.Remove(MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME)
        Next
        BeforeRowsDeleted = Nothing
    End Sub

    Private Sub grdBETCONV3_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdBETCONV3.BeforeRowsDeleted
        BeforeRowsDeleted = New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim MYS_SCHEMA_NAME As String = grow.Cells("MYS_SCHEMA_NAME").Value
            Dim MYS_TABLE_NAME As String = grow.Cells("MYS_TABLE_NAME").Value
            BeforeRowsDeleted.Add(MYS_SCHEMA_NAME & "." & MYS_TABLE_NAME)
        Next
    End Sub

    Private Sub grdBETCONV3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdBETCONV3.InitializeRow
        For Each COLUMN_NAME As String In ORA_TBL_SCHEMA.Keys
            If COLUMN_NAME = "PROGRAM_CODE" Or COLUMN_NAME = "ID" Then
            Else
                Dim Cx As Integer = ORA_TBL_SCHEMA(COLUMN_NAME)
                Dim C As String = "COL_" & Format(Cx, "000")
                If e.Row.Cells(C).Value & "" = COLUMN_NAME Then
                    e.Row.Cells(C).Appearance.BackColor = Drawing.Color.Empty
                Else
                    e.Row.Cells(C).Appearance.BackColor = Drawing.Color.Yellow
                End If
            End If
        Next
    End Sub

    Private Sub grdBETCONVS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdBETCONVS.AfterRowActivate
        Dim COLUMN_NAME As String = grdBETCONVS.ActiveRow.Cells("COLUMN_NAME").Value
        If Not ORA_TBL_SCHEMA.ContainsKey(COLUMN_NAME.ToUpper) Then
            grdBETCONVS.ActiveRow.Cells("COLUMN_NAME").Appearance.BackColor = Drawing.Color.Yellow
        Else
            grdBETCONVS.ActiveRow.Cells("COLUMN_NAME").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdBETCONVD_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdBETCONVD.InitializeRow
        If e.Row.Band.Key = "BETCONVD" Then
            Dim DEFINITION As String = e.Row.Cells("DEFINITION").Value & ""
            For Each SCHEMA_NAME As String In SCHEMAs.Keys
                Dim RC As String = e.Row.Cells(SCHEMA_NAME).Value & ""
                If RC <> DEFINITION Then
                    If RC = "" Then
                        e.Row.Cells(SCHEMA_NAME).Appearance.BackColor = Drawing.Color.Yellow
                    Else
                        e.Row.Cells(SCHEMA_NAME).Appearance.ForeColor = Drawing.Color.Red
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub tabTables_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabTables.SelectedTabChanged
        Setup_tabTables()
    End Sub

    Sub Setup_tabTables()
        UltraExplorerBar1.Groups("Conversion List").Visible = (tabTables.SelectedTab.Key = "mySQL Tables - List")
    End Sub

    Sub Setup_BETCONV0()
        Dim dvw As DataView = DirectCast(grdBETCONV0.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If chkShowAll.Checked Then
        Else
            sql = "CONVERT= '1'"
        End If
        dvw.RowFilter = sql
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowAll.CheckedChanged
        Setup_BETCONV0()
    End Sub

    Private Sub grdBETCONV0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdBETCONV0.InitializeRow
        If e.Row.Band.Key = "BETCONV0" Then
            If e.Row.Cells("CONVERT").Value & "" = "1" Then
                e.Row.Cells("CONVERT").Appearance.BackColor = Drawing.Color.Empty
            Else
                e.Row.Cells("CONVERT").Appearance.BackColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Public Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        Dim COLUMN_NAME As String = summarySettings.Key
        CustomValue = 0

        Select Case grd.Name
            Case "grdBETTABLX"
                CustomValue = 0
                For Each grow As UltraWinGrid.UltraGridRow In rows
                    If grow.Cells(COLUMN_NAME).Value & "" <> "" Then
                        CustomValue += 1
                    End If
                Next
                Return CustomValue
            Case "grdBETCONVD"
                CustomValue = 0
                For Each grow As UltraWinGrid.UltraGridRow In rows
                    If grow.Cells(COLUMN_NAME).Value & "" <> grow.Cells("DEFINITION").Value & "" Then
                        CustomValue += 1
                    End If
                Next

                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Issues Only"), UltraWinToolbars.StateButtonTool)
                grdBETCONVD.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = (tlb_sbt.Checked And CustomValue = 0)
                grdBETCONVD.DisplayLayout.Bands(1).Columns(COLUMN_NAME).Hidden = (tlb_sbt.Checked And CustomValue = 0)

                Return CustomValue
        End Select
    End Function

    Private Sub cmdList_Click(sender As System.Object, e As System.EventArgs) Handles cmdList.Click
        Dim sqlw As String = "TABLE_NAME = '" & lblTableName.Text & "'"
        If optAllOnly.Value = "A" Then
        Else
            sqlw &= " and SCHEMA_NAME = '" & lblSchemaName.Text & "'"
        End If

        For Each rowBETCONV0 As DataRow In dst.Tables("BETCONV0").Select(sqlw)
            rowBETCONV0.Item("CONVERT") = IIf(optAddRemove.Value = "A", "1", "0")
        Next
        Update_Record_TDA("BETCONV0")

        If grdBETCONV0.ActiveRow Is Nothing Then
            UltraExplorerBar1.Groups("Conversion List").Visible = False
        End If

        grdBETCONV0.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Private Sub optAddRemove_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAddRemove.ValueChanged
        setup_cmdList_Caption()
    End Sub

    Private Sub optAllOnly_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAllOnly.ValueChanged
        setup_cmdList_Caption()
    End Sub

    Sub setup_cmdList_Caption()
        cmdList.Text = ""
        If optAddRemove.Value = "A" Then
            cmdList.Text &= "Add Table " & lblTableName.Text
        Else
            cmdList.Text &= "Remove Table " & lblTableName.Text
        End If
        cmdList.Text &= vbCrLf & " from the Conversion List"
        If optAllOnly.Value = "A" Then
            cmdList.Text &= vbCrLf & "For All Schemas"
        Else
            cmdList.Text &= vbCrLf & "For " & lblSchemaName.Text & " Only"
        End If
    End Sub

    Sub Get_Data(TABLE_NAME As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Getting Sample Data")

        Dim rowBETTABLX As DataRow = dst.Tables("BETTABLX").Rows.Find(TABLE_NAME)

        For Each SCHEMA_NAME As String In SCHEMAs.Keys
            If rowBETTABLX.Item(SCHEMA_NAME) & "" <> "" Then
                ASCMAIN1.Progress("-", SCHEMA_NAME)
                Dim tblT As DataTable = Fill_Table_from_MYSQL(Nothing, "Select * from " & SCHEMA_NAME & "." & TABLE_NAME)
                For Each DC As DataColumn In tblT.Columns
                    Dim ROW_NO As Integer = 0
                    For Each row As DataRow In ASCDATA1.SelectDistinct(tblT, New String() {DC.ColumnName}).Rows
                        ROW_NO += 1
                        Dim rowBETCONVT As DataRow = dst.Tables("BETCONVT").Rows.Find _
                            (New Object() {TABLE_NAME, DC.ColumnName, ROW_NO})
                        If rowBETCONVT Is Nothing Then
                            rowBETCONVT = dst.Tables("BETCONVT").Rows.Add _
                                (New Object() {TABLE_NAME, DC.ColumnName, ROW_NO})
                        End If
                        rowBETCONVT.Item(SCHEMA_NAME) = row.Item(DC.ColumnName)
                        If ROW_NO = 10 Then Exit For
                    Next
                Next
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Generate_Oracle_DDL(TABLE_NAME As String)

        Dim KEY As String = ""
        Dim SQL As String = ""
        For Each row As DataRow In dst.Tables("BETCONVD").Select("TABLE_NAME = '" & TABLE_NAME & "'")

            Dim DEFINITION As String = row.Item("DEFINITION")
            Dim ORA_COL_DEF As String = ""
            If DEFINITION Like "System.String*" Then
                ORA_COL_DEF = Replace(DEFINITION, "System.String", "VARCHAR2")
            ElseIf DEFINITION = "MySql.Data.Types.MySqlDateTime" Then
                ORA_COL_DEF = "DATE"
            ElseIf DEFINITION = "System.Int32" Or DEFINITION = "System.UInt32" Then
                ORA_COL_DEF = "NUMBER (" & "10" & ",0)"
            ElseIf DEFINITION = "System.Byte[]" Then
                ORA_COL_DEF = "VARCHAR2(20)"
            ElseIf DEFINITION Like "System.Single*" Then
                ORA_COL_DEF = "NUMBER(20,6)"
            ElseIf DEFINITION = "System.TimeSpan" Then
                ORA_COL_DEF = "DATE"
            Else
                Stop

            End If
            Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
            If COLUMN_NAME.Length > 30 Then
                COLUMN_NAME = Mid(COLUMN_NAME, 1, 30)
            End If
            ORA_COL_DEF = "," & vbCrLf & COLUMN_NAME & " " & ORA_COL_DEF

            If COLUMN_NAME.ToUpper = "ID" Then
                SQL = ORA_COL_DEF & SQL
                KEY &= ",ID"
            Else
                SQL &= ORA_COL_DEF
            End If
        Next
        If KEY <> "" Then
            KEY = "PROGRAM_CODE" & KEY
        End If
        SQL = "Create Table BETXXXXX (PROGRAM_CODE VARCHAR2(20)" & SQL & IIf(KEY <> "", "," & vbCrLf & "Primary Key (" & KEY & ")", "") & ");"
        Using F As New ASFMSGBF
             F.Get_txtblock_from_User("DDL for Oracle Table", "Oracle DDL", SQL)
        End Using
    End Sub
End Class