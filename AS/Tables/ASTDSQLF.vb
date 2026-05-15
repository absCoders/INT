Imports Infragistics.Win.UltraWinGrid

Public Class ASTDSQLF

    Dim COLUMN_NAME_del As String
    Dim COLUMN_NAME_old As String
    Dim COLUMN_NAME_new As String

    Dim tblCOLUMN_NAMEs As New DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ASTDSQLA", "*", 1)

            Create_TDA(.Tables.Add, "ASTDSQLB", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLJ", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLC", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLD", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLH", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLS", "*", 1)

            Create_TDA(.Tables.Add, "ASTDSQLK", "*", 0, False)

            Create_TDA(.Tables.Add, "ASTDSQLV", "*", 1)
            Create_TDA(.Tables.Add, "ASTDSQLW", "*", 1)

            .Relations.Add("ASTDSQLJ",
            New DataColumn() { .Tables("ASTDSQLB").Columns("FORM_NAME"), .Tables("ASTDSQLB").Columns("DATA_SOURCE")},
            New DataColumn() { .Tables("ASTDSQLJ").Columns("FORM_NAME"), .Tables("ASTDSQLJ").Columns("DATA_SOURCE")})

            .Relations.Add("ASTDSQLD",
            New DataColumn() { .Tables("ASTDSQLJ").Columns("FORM_NAME"), .Tables("ASTDSQLJ").Columns("DATA_SOURCE"), .Tables("ASTDSQLJ").Columns("TABLE_NAME")},
            New DataColumn() { .Tables("ASTDSQLD").Columns("FORM_NAME"), .Tables("ASTDSQLD").Columns("DATA_SOURCE"), .Tables("ASTDSQLD").Columns("TABLE_NAME")})

            .Relations.Add("ASTDSQLC",
            New DataColumn() { .Tables("ASTDSQLB").Columns("FORM_NAME"), .Tables("ASTDSQLB").Columns("DATA_SOURCE")},
            New DataColumn() { .Tables("ASTDSQLC").Columns("FORM_NAME"), .Tables("ASTDSQLC").Columns("DATA_SOURCE")})

            .Relations.Add("ASTDSQLW",
            New DataColumn() { .Tables("ASTDSQLV").Columns("FORM_NAME"), .Tables("ASTDSQLV").Columns("VALUE_LIST_NAME")},
            New DataColumn() { .Tables("ASTDSQLW").Columns("FORM_NAME"), .Tables("ASTDSQLW").Columns("VALUE_LIST_NAME")})

        End With

        grdASTDSQLA.DataSource = dst.Tables("ASTDSQLA")
        grdASTDSQLB.DataSource = dst.Tables("ASTDSQLB")
        grdASTDSQLH.DataSource = dst.Tables("ASTDSQLH")
        grdASTDSQLS.DataSource = dst.Tables("ASTDSQLS")
        grdASTDSQLV.DataSource = dst.Tables("ASTDSQLV")

        Call Fill_Records("ASTDSQLK")

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & "" & "'"
        Else
            ASCMAIN1.sql = "Select COLUMN_NAME from USER_TAB_COLUMNS where TABLE_NAME = '" & "" & "'"
        End If
        tblCOLUMN_NAMEs = ASCDATA1.GetDataTable()
        grdCOLUMN_NAMEs.DataSource = tblCOLUMN_NAMEs
        grdCOLUMN_NAMEs.DisplayLayout.Bands(0).SortedColumns.Add(grdCOLUMN_NAMEs.DisplayLayout.Bands(0).Columns("COLUMN_NAME"), False)

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select TABLE_NAME from INFORMATION_SCHEMA.TABLES"
        Else
            ASCMAIN1.sql = "Select TABLE_NAME from USER_TABLES"
        End If
        cmbTable.DataSource = ASCDATA1.GetDataTable("", "ASTTABD1")
        cmbTable.DisplayLayout.Bands(0).SortedColumns.Add(cmbTable.DisplayLayout.Bands(0).Columns("TABLE_NAME"), False)

    End Sub

    Private Sub grdASTDSQLA_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTDSQLA.AfterRowsDeleted
        For Each rowASTDSQLC As DataRow In dst.Tables("ASTDSQLC").Select("COLUMN_NAME = '" & COLUMN_NAME_del & "'")
            rowASTDSQLC.Delete()
        Next
    End Sub

    Private Sub grdASTDSQLA_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTDSQLA.AfterRowUpdate
        If COLUMN_NAME_new <> "" Or COLUMN_NAME_old <> "" Then
            If COLUMN_NAME_old = "" Then
                Call Add_to_ASTDSQLC()
            Else
                For Each rowASTDSQLC As DataRow In dst.Tables("ASTDSQLC").Select("COLUMN_NAME = '" & COLUMN_NAME_old & "'")
                    rowASTDSQLC.Item("COLUMN_NAME") = COLUMN_NAME_new
                Next
            End If
        End If
        Me.MdiParent = ASCMAIN1.MainForm
        'Me.MdiParent = Nothing
    End Sub

    Sub Add_to_ASTDSQLC()
        For Each rowASTDSQLB As DataRow In dst.Tables("ASTDSQLB").Select
            Dim rowASTDSQLC As DataRow = dst.Tables("ASTDSQLC").NewRow
            rowASTDSQLC.Item("FORM_NAME") = rowASTDSQLB.Item("FORM_NAME")
            rowASTDSQLC.Item("DATA_SOURCE") = rowASTDSQLB.Item("DATA_SOURCE")
            rowASTDSQLC.Item("COLUMN_NAME") = COLUMN_NAME_new
            dst.Tables("ASTDSQLC").Rows.Add(rowASTDSQLC)
        Next
    End Sub

    Private Sub grdASTDSQLA_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTDSQLA.BeforeRowsDeleted
        COLUMN_NAME_del = e.Rows(0).Cells("COLUMN_NAME").Text
    End Sub

    Private Sub grdASTDSQLA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTDSQLA.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            If .ActiveRow.Cells("FORM_NAME").Text = "" Then
                .ActiveRow.Cells("FORM_NAME").Value = Me.Absx1.txtFor("FORM_NAME").Text
                .ActiveRow.Cells("SORTABLE").Value = Yes_if_any_are_SORTABLE()
            End If
            If .ActiveRow.Cells("COLUMN_CAPTION").Text.Trim = "" Then
                Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(.ActiveRow.Cells("COLUMN_NAME").Value & "")
                If Not rowASTDSQLK Is Nothing Then
                    .ActiveRow.Cells("COLUMN_CAPTION").Value = rowASTDSQLK.Item("COLUMN_CAPTION")
                End If
            End If

            COLUMN_NAME_new = ""
            COLUMN_NAME_old = ""
            If e.Row.Cells("COLUMN_NAME").DataChanged Then
                COLUMN_NAME_new = e.Row.Cells("COLUMN_NAME").Text & ""
                If Not e.Row.IsAddRow Then
                    COLUMN_NAME_old = e.Row.Cells("COLUMN_NAME").OriginalValue & ""
                End If
                If COLUMN_NAME_new = COLUMN_NAME_old Then
                    COLUMN_NAME_new = ""
                    COLUMN_NAME_old = ""
                End If
            End If
        End With
    End Sub

    Private Sub grdASTDSQLB_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTDSQLB.AfterRowUpdate
        If e.Row.Band.Key = "ASTDSQLB" Then ' And e.Row.IsAddRow Then
            For Each rowASTDSQLA As DataRow In dst.Tables("ASTDSQLA").Rows
                Try
                    If dst.Tables("ASTDSQLC").Rows.Find _
                    (New String() {
                     rowASTDSQLA.Item("FORM_NAME"),
                     e.Row.Cells("DATA_SOURCE").Text,
                     rowASTDSQLA.Item("COLUMN_NAME")}) Is Nothing Then
                        Dim rowASTDSQLC As DataRow = dst.Tables("ASTDSQLC").NewRow
                        rowASTDSQLC.Item("FORM_NAME") = rowASTDSQLA.Item("FORM_NAME")
                        rowASTDSQLC.Item("DATA_SOURCE") = e.Row.Cells("DATA_SOURCE").Text
                        rowASTDSQLC.Item("COLUMN_NAME") = rowASTDSQLA.Item("COLUMN_NAME")
                        dst.Tables("ASTDSQLC").Rows.Add(rowASTDSQLC)
                    End If

                Catch ex As Exception

                End Try
            Next
        End If

        If e.Row.Band.Key = "ASTDSQLC" Then ' And e.Row.IsAddRow Then
            If e.Row.Cells("TABLE_NAME").Text <> "" Then
                Try
                    If dst.Tables("ASTDSQLJ").Rows.Find _
                    (New String() {
                     e.Row.Cells("FORM_NAME").Text,
                     e.Row.Cells("DATA_SOURCE").Text,
                     e.Row.Cells("TABLE_NAME").Text}) Is Nothing Then
                        Dim rowASTDSQLJ As DataRow = dst.Tables("ASTDSQLJ").NewRow
                        rowASTDSQLJ.Item("FORM_NAME") = e.Row.Cells("FORM_NAME").Text
                        rowASTDSQLJ.Item("DATA_SOURCE") = e.Row.Cells("DATA_SOURCE").Text
                        rowASTDSQLJ.Item("TABLE_NAME") = e.Row.Cells("TABLE_NAME").Text
                        dst.Tables("ASTDSQLJ").Rows.Add(rowASTDSQLJ)
                    End If
                Catch ex As Exception

                End Try
            End If
        End If

    End Sub

    Private Sub grdASTDSQLB_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTDSQLB.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            If e.Row.Band.Key = "ASTDSQLB" Then
                If .ActiveRow.Cells("FORM_NAME").Text = "" Then
                    .ActiveRow.Cells("FORM_NAME").Value = Absx1.txtFor("FORM_NAME").Text
                End If
            End If
        End With
    End Sub

    Sub Add_Key_Column(ByVal COLUMN_NAME As String)
        Dim rowASTDSQLA As DataRow = dst.Tables("ASTDSQLA").NewRow
        rowASTDSQLA.Item("FORM_NAME") = Absx1.CtlFor("FORM_NAME").Text
        rowASTDSQLA.Item("COLUMN_NAME") = COLUMN_NAME
        Dim R As DataRow = dst.Tables("ASTDSQLK").Rows.Find(COLUMN_NAME)
        If Not R Is Nothing Then
            rowASTDSQLA.Item("COLUMN_CAPTION") = R.Item("COLUMN_CAPTION")
        Else
            rowASTDSQLA.Item("COLUMN_CAPTION") = ASCMAIN1.Make_Caption(COLUMN_NAME)
        End If

        rowASTDSQLA.Item("SORTABLE") = Yes_if_any_are_SORTABLE()
        dst.Tables("ASTDSQLA").Rows.Add(rowASTDSQLA)

        COLUMN_NAME_new = COLUMN_NAME
        Call Add_to_ASTDSQLC()
    End Sub

    Function Yes_if_any_are_SORTABLE() As String
        If dst.Tables("ASTDSQLA").Select("SORTABLE = '1'").Length = 0 Then
            Return "0"
        Else
            Return "1"
        End If
    End Function

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdASTDSQLS, "B", "Re-Number from Here")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Re-Number from Here"

                Dim COLUMN_SEQ As Integer = Val(grd.ActiveRow.Cells("COLUMN_SEQ").Value & "")

                For Each rowASTDSQLS As DataRow In dst.Tables("ASTDSQLS").Select($"COLUMN_SEQ >= {CStr(COLUMN_SEQ)}", "COLUMN_SEQ DESC")
                    rowASTDSQLS.Item("COLUMN_SEQ") = Val(rowASTDSQLS.Item("COLUMN_SEQ") & "") + 1
                Next

        End Select

    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        For Each rowASTDSQLB As DataRow In dst.Tables("ASTDSQLB").Select("", "")
            For Each rowASTDSQLC As DataRow In dst.Tables("ASTDSQLC").Select("FORM_NAME = '" & rowASTDSQLB.Item("FORM_NAME") & "' and DATA_SOURCE = '" & rowASTDSQLB.Item("DATA_SOURCE") & "'")
                If rowASTDSQLC.Item("TABLE_NAME") & "" = rowASTDSQLB.Item("TABLE_NAME") & "" Then
                    rowASTDSQLC.Item("TABLE_NAME") = ""
                End If
                If rowASTDSQLC.Item("TABLE_NAME") & "" = "" And rowASTDSQLC.Item("COLUMN_EXPRESSION") & "" = "" And rowASTDSQLC.Item("NO_FILTER") & "" <> "1" Then
                    rowASTDSQLC.Delete()
                End If
            Next
            For Each rowASTDSQLD As DataRow In TBLs("ASTDSQLD").Select("FORM_NAME = '" & rowASTDSQLB.Item("FORM_NAME") & "' and DATA_SOURCE = '" & rowASTDSQLB.Item("DATA_SOURCE") & "'")
                If rowASTDSQLD.Item("TABLE_NAME_JOIN") & "" = rowASTDSQLB.Item("TABLE_NAME") & "" Then
                    rowASTDSQLD.Item("TABLE_NAME_JOIN") = ""
                End If
            Next
        Next

        For Each rowASTDSQLA As DataRow In dst.Tables("ASTDSQLA").Select("", "")
            Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(rowASTDSQLA.Item("COLUMN_NAME"))
            Dim COLUMN_CAPTION_default As String = ""
            If rowASTDSQLK IsNot Nothing Then
                COLUMN_CAPTION_default = rowASTDSQLK.Item("COLUMN_CAPTION")
            End If
            If rowASTDSQLA.Item("COLUMN_CAPTION") & "" = COLUMN_CAPTION_default Then
                rowASTDSQLA.Item("COLUMN_CAPTION") = DBNull.Value
            End If
        Next

        Dim sql As String = ""

        sql = "Delete from ASTDSQLD where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
        Update_Record_TDA("ASTDSQLD", sql)

        sql = "Delete from ASTDSQLA where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
        Update_Record_TDA("ASTDSQLA", sql)
        Update_Record_TDA("ASTDSQLB")
        Update_Record_TDA("ASTDSQLC")
        Update_Record_TDA("ASTDSQLJ")
        Update_Record_TDA("ASTDSQLH")
        sql = "Delete from ASTDSQLS where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
        Update_Record_TDA("ASTDSQLS", sql)
        Update_Record_TDA("ASTDSQLV")
        Update_Record_TDA("ASTDSQLW")

    End Sub

    Overrides Sub Show_Record_Special()
        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = Absx1.txtFor("FORM_NAME")
        Call Clear_Record_Special()
        Call Load_Report_Form(txtctl.Text)
        cmbCopyFrom.Enabled = True
        cmbCopyFrom.ReadOnly = False

    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

        Call Fill_Records("ASTDSQLA", FORM_NAME)
        For Each r As DataRow In dst.Tables("ASTDSQLA").Rows
            If r.Item("COLUMN_CAPTION") & "" = "" Then
                Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(r.Item("COLUMN_NAME"))
                If Not rowASTDSQLK Is Nothing Then
                    r.Item("COLUMN_CAPTION") = rowASTDSQLK.Item("COLUMN_CAPTION")
                End If
            End If
        Next

        EnforceConstraints(False)

        Call Fill_Records("ASTDSQLB", FORM_NAME)
        Call Fill_Records("ASTDSQLC", FORM_NAME)

        For Each rowASTDSQLC As DataRow In dst.Tables("ASTDSQLC").Rows
            rowASTDSQLC.Delete()
        Next

        ASCMAIN1.sql = "SELECT X.FORM_NAME, X.DATA_SOURCE, X.COLUMN_NAME, " _
            & "ASTDSQLC.TABLE_NAME, ASTDSQLC.COLUMN_EXPRESSION, ASTDSQLC.JOIN_SPECIAL, ASTDSQLC.NO_FILTER " _
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
        Dim tblx As DataTable = ASCDATA1.GetDataTable

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            dst.Tables("ASTDSQLC").LoadDataRow(row.ItemArray, False)
        Next


        Call Fill_Records("ASTDSQLD", FORM_NAME)
        Call Fill_Records("ASTDSQLJ", FORM_NAME)

        If dst.Tables("ASTDSQLB").Rows.Find(New String() {FORM_NAME, "*"}) Is Nothing Then
            Dim row As DataRow = dst.Tables("ASTDSQLB").NewRow
            row.Item("FORM_NAME") = FORM_NAME
            row.Item("DATA_SOURCE") = "*"
            dst.Tables("ASTDSQLB").Rows.Add(row)
        End If

        Call Fill_Records("ASTDSQLH", FORM_NAME)
        Call Fill_Records("ASTDSQLS", FORM_NAME)
        Call Fill_Records("ASTDSQLV", FORM_NAME)
        Call Fill_Records("ASTDSQLW", FORM_NAME)

        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            dst.EnforceConstraints = False
            dst.Tables("ASTDSQLA").Rows.Clear()
            dst.Tables("ASTDSQLB").Rows.Clear()
            dst.Tables("ASTDSQLC").Rows.Clear()
            dst.Tables("ASTDSQLJ").Rows.Clear()
            dst.Tables("ASTDSQLD").Rows.Clear()
            dst.Tables("ASTDSQLH").Rows.Clear()
            dst.Tables("ASTDSQLS").Rows.Clear()
            dst.Tables("ASTDSQLV").Rows.Clear()
            dst.Tables("ASTDSQLW").Rows.Clear()
            dst.EnforceConstraints = True
        End If

        tblCOLUMN_NAMEs.Rows.Clear()

        cmbCopyFrom.Enabled = True
        cmbCopyFrom.ReadOnly = True

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            cmbTable.Text = ""
            cmbCopyFrom.Text = ""
            UltraTabControl1.Tabs(0).Selected = True
        End If
        cmbTable.Enabled = tf
        'cmbCopyFrom.Enabled = tf
        'cmbCopyFrom.ReadOnly = tf
        'TabControl1.Visible = tf
        grdASTDSQLA.Enabled = tf
        grdASTDSQLB.Enabled = tf
        grdASTDSQLH.Enabled = tf
        grdASTDSQLS.Enabled = tf

        grdASTDSQLV.Enabled = tf

        grdCOLUMN_NAMEs.Enabled = tf

        cmdSchema.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        txtSCHEMA.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        lblLINK.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        txtLINK.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        Set_Read_Only_for_ctl(txtSCHEMA, Not txtSCHEMA.Visible)
        Set_Read_Only_for_ctl(txtLINK, Not txtSCHEMA.Visible)
        ' Set_Read_Only_for_ctl(cmbCopyFrom, tf)
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                If dst.Tables("ASTDSQLA").Select("COLUMN_LAST = '1'").Length > 1 Then
                    EMsg = EMsg & vbCr & "Only 1 Column may be Selected to be Last"
                End If

                If dst.Tables("ASTDSQLA").Select("SORTABLE = '1'").Length > 0 Then
                    For Each rowASTDSQLB As DataRow In dst.Tables("ASTDSQLB").Select("TABLE_NAME is Null or TABLE_NAME = ''")
                        EMsg = EMsg & vbCr & "No Table Name defined for Data Source " & rowASTDSQLB.Item("DATA_SOURCE")
                    Next
                End If

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTDSQLC"), New String() {"FORM_NAME", "DATA_SOURCE", "TABLE_NAME"}).Rows
                    Dim rowASTDSQLB As DataRow = dst.Tables("ASTDSQLB").Rows.Find(New String() {row.Item("FORM_NAME"), row.Item("DATA_SOURCE")})
                    If row.Item("TABLE_NAME") & "" = "" Or row.Item("TABLE_NAME") & "" = rowASTDSQLB.Item("TABLE_NAME") & "" Then
                    Else
                        Dim foundrow As DataRow = dst.Tables("ASTDSQLJ").Rows.Find(New String() {row.Item("FORM_NAME"), row.Item("DATA_SOURCE"), row.Item("TABLE_NAME")})
                        If foundrow Is Nothing Then
                            EMsg = EMsg & vbCr & "No Join Record set up for Table " & row.Item("TABLE_NAME")
                        End If
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(TBLs("ASTDSQLD"), New String() {"FORM_NAME", "DATA_SOURCE", "TABLE_NAME", "TABLE_NAME_JOIN"}).Rows
                    Dim rowASTDSQLB As DataRow = dst.Tables("ASTDSQLB").Rows.Find(New String() {row.Item("FORM_NAME"), row.Item("DATA_SOURCE")})
                    If row.Item("TABLE_NAME_JOIN") & "" = "" Or row.Item("TABLE_NAME_JOIN") & "" = rowASTDSQLB.Item("TABLE_NAME") & "" Then
                    Else
                        Dim foundrow As DataRow = dst.Tables("ASTDSQLJ").Rows.Find(New String() {row.Item("FORM_NAME"), row.Item("DATA_SOURCE"), row.Item("TABLE_NAME_JOIN")})
                        If foundrow Is Nothing Then
                            EMsg = EMsg & vbCr & "No Join Record set up for Table " & row.Item("TABLE_NAME_JOIN")
                        End If
                    End If
                Next

                If Absx1.CtlFor("FORM_CAPTION").Text = "" Then
                    EMsg = EMsg & vbCr & "No Form Caption for the Report Definition"
                End If
        End Select

    End Sub
#End Region

    Private Sub grdASTDSQLA_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTDSQLA.InitializeLayout
        grdASTDSQLA.DisplayLayout.Bands(0).Columns("SORTABLE").Editor.DataFilter = New CheckEditorDataFilter
    End Sub

    Private Sub grdASTDSQLB_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTDSQLB.InitializeLayout
        grdASTDSQLB.DisplayLayout.Bands("ASTDSQLJ").Columns("OUTER_JOIN").Editor.DataFilter = New CheckEditorDataFilter
        grdASTDSQLB.DisplayLayout.Bands("ASTDSQLJ").Columns("ALWAYS_JOIN").Editor.DataFilter = New CheckEditorDataFilter
        grdASTDSQLB.DisplayLayout.Bands("ASTDSQLC").Columns("JOIN_SPECIAL").Editor.DataFilter = New CheckEditorDataFilter
    End Sub

    Private Sub cmbCopyFrom_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCopyFrom.AfterCloseUp
        If ScreenMode Then
            If cmbCopyFrom.Text = "" Then
                Exit Sub
            End If

            If dst.Tables("ASTDSQLA").Select("", "", DataViewRowState.CurrentRows).Length <> 0 Then
                If chkAppend.Checked Then
                    If MsgBox("Append to Current Definition with Information Copied from Report " & cmbCopyFrom.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                Else
                    If MsgBox("Replace Current Definition with Information Copied from Report " & cmbCopyFrom.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If


            Copy_Definition(cmbCopyFrom.Text)

        End If

    End Sub

    Sub Copy_Definition(ByVal FORM_NAME_COPY_FROM As String)

        If chkAppend.Checked Then
            ASCMAIN1.sql = "Select * from ASTDSQLA where FORM_NAME = '" & FORM_NAME_COPY_FROM & "'"
            For Each rowASTDSQLA As DataRow In ASCDATA1.GetDataTable.Rows
                COLUMN_NAME = rowASTDSQLA.Item("COLUMN_NAME")
                If dst.Tables("ASTDSQLA").Select("COLUMN_NAME = '" & COLUMN_NAME & "'").GetLength(0) = 0 Then
                    Call Add_Key_Column(COLUMN_NAME)
                End If
            Next
        Else
            Clear_Record_Special()
            Call Load_Report_Form(FORM_NAME_COPY_FROM)
            For Each tbl As DataTable In New DataTable() {dst.Tables("ASTDSQLA"), dst.Tables("ASTDSQLB"), dst.Tables("ASTDSQLC"), dst.Tables("ASTDSQLD"), dst.Tables("ASTDSQLJ"), dst.Tables("ASTDSQLH"), dst.Tables("ASTDSQLS"), dst.Tables("ASTDSQLV"), dst.Tables("ASTDSQLW")}
                If tbl.TableName = "ASTDSQLA" Or tbl.TableName = "ASTDSQLB" Or tbl.TableName = "ASTDSQLH" Or tbl.TableName = "ASTDSQLS" Or tbl.TableName = "ASTDSQLV" Then
                    For Each row As DataRow In tbl.Rows
                        row.Item("FORM_NAME") = Absx1.txtFor("FORM_NAME").Text
                    Next
                End If

                tbl.AcceptChanges()
                For i As Integer = 0 To tbl.Rows.Count - 1
                    tbl.Rows(i).SetAdded()
                Next
            Next
        End If
    End Sub

    Private Sub cmbTable_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTable.AfterCloseUp
        If cmbTable.Text <> "" Then
            tblCOLUMN_NAMEs.Rows.Clear()
            ASCMAIN1.sql = "Select COLUMN_NAME from USER_TAB_COLUMNS where TABLE_NAME = '" & cmbTable.Text & "'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable()
            If tbl.Rows.Count > 0 Then
                For I As Integer = 0 To tbl.Rows.Count - 1
                    tblCOLUMN_NAMEs.ImportRow(tbl.Rows(I))
                Next
            End If

            'tblCOLUMN_NAMEs = ASCDATA1.GetDataTable()
            'grdCOLUMN_NAMEs.DataBind()
            'grdCOLUMN_NAMEs.DataSource = tblCOLUMN_NAMEs
            'grdCOLUMN_NAMEs.DisplayLayout.Bands(0).SortedColumns.Add(grdCOLUMN_NAMEs.DisplayLayout.Bands(0).Columns("COLUMN_NAME"), False)

        End If
    End Sub

    Private Sub grdCOLUMN_NAMEs_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdCOLUMN_NAMEs.AfterSelectChange
        ' Stop
    End Sub

    Private Sub grdCOLUMN_NAMEs_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdCOLUMN_NAMEs.MouseUp
        Dim x As UIElement
        x = grdCOLUMN_NAMEs.DisplayLayout.UIElement.ElementFromPoint(New System.Drawing.Point(e.X, e.Y))
        If TypeOf x Is UltraWinGrid.RowSelectorUIElement Then
            'Dim r As UltraWinGrid.RowSelectorUIElement
            'r = DirectCast(x, UltraWinGrid.RowSelectorUIElement)
            'Stop
            If grdCOLUMN_NAMEs.Selected.Rows.Count = 1 Then
                COLUMN_NAME = grdCOLUMN_NAMEs.ActiveRow.Cells("COLUMN_NAME").Text
                If dst.Tables("ASTDSQLA").Select("COLUMN_NAME = '" & COLUMN_NAME & "'").GetLength(0) = 0 Then
                    Call Add_Key_Column(COLUMN_NAME)
                End If
            End If

        End If
    End Sub

    Private Sub grdASTDSQLA_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTDSQLA.InitializeRow
        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text
        Dim COLUMN_CAPTION As String = e.Row.Cells("COLUMN_CAPTION").Value & ""
        Dim COLUMN_CAPTION_default As String = ""
        Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(COLUMN_NAME)
        If Not rowASTDSQLK Is Nothing Then
            COLUMN_CAPTION_default = rowASTDSQLK.Item("COLUMN_CAPTION")
        End If
        If COLUMN_CAPTION = COLUMN_CAPTION_default Then
            e.Row.Cells("COLUMN_CAPTION").Appearance.ForeColor = Drawing.Color.Black
        Else
            If Not rowASTDSQLK Is Nothing Then
                e.Row.Cells("COLUMN_CAPTION").Appearance.ForeColor = Drawing.Color.Red
            Else
                If COLUMN_CAPTION <> "" Then
                    Try
                        ' don't know why this insert did not work 
                        ASCMAIN1.sql = "Insert into ASTDSQLK (COLUMN_NAME,COLUMN_CAPTION) " _
                        & " VALUES ('" & COLUMN_NAME & "','" & COLUMN_CAPTION & "')"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                        rowASTDSQLK = dst.Tables("ASTDSQLK").NewRow
                        rowASTDSQLK.Item("COLUMN_NAME") = COLUMN_NAME
                        rowASTDSQLK.Item("COLUMN_CAPTION") = COLUMN_CAPTION
                        dst.Tables("ASTDSQLK").Rows.Add(rowASTDSQLK)

                        e.Row.Cells("COLUMN_CAPTION").Appearance.ForeColor = Drawing.Color.Black
                        e.Row.Description = "Column Caption '" & COLUMN_CAPTION & "' was added to the Default Captions Table"
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub grdASTDSQLB_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdASTDSQLB.KeyDown
        'Call ASCMAIN1.Navigate_like_Excel(grdASTDSQLB, e)
    End Sub

    Private Sub grdASTDSQLA_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdASTDSQLA.KeyDown
        'Call ASCMAIN1.Navigate_like_Excel(grdASTDSQLA, e)
    End Sub

    Private Sub grdASTDSQLH_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTDSQLH.BeforeRowUpdate
        With grdASTDSQLH.ActiveRow
            If .Cells("FORM_NAME").Text = "" Then
                .Cells("FORM_NAME").Value = Absx1.txtFor("FORM_NAME").Text
                If Val(.Cells("COLUMN_SEQ").Value & "") = 0 Then
                    Dim COLUMN_SEQ As Integer = Val(dst.Tables("ASTDSQLH").Compute("MAX(COLUMN_SEQ)", "") & "")
                    .Cells("COLUMN_SEQ").Value = COLUMN_SEQ + 1
                    If COLUMN_SEQ <> 0 Then
                        Dim row As DataRow = dst.Tables("ASTDSQLH").Rows.Find(New Object() {Absx1.txtFor("FORM_NAME").Text, COLUMN_SEQ})
                        If .Cells("TABLE_NAME").Value & "" = "" Then
                            .Cells("TABLE_NAME").Value = row.Item("TABLE_NAME")
                        End If
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grdASTDSQLS_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTDSQLS.BeforeRowUpdate
        With grdASTDSQLS.ActiveRow
            If .Cells("FORM_NAME").Text = "" Then
                .Cells("FORM_NAME").Value = Absx1.txtFor("FORM_NAME").Text
                If Val(.Cells("COLUMN_SEQ").Value & "") = 0 Then
                    Dim COLUMN_SEQ_max As Integer = Val(dst.Tables("ASTDSQLS").Compute("MAX(COLUMN_SEQ)", "") & "")
                    Dim COLUMN_SEQ As Integer = COLUMN_SEQ_max
                    If COLUMN_SEQ_max > 1 Then
                        For i As Integer = 1 To COLUMN_SEQ_max - 1
                            If dst.Tables("ASTDSQLS").Rows.Find(New Object() {Absx1.txtFor("FORM_NAME").Text, i}) Is Nothing Then
                                COLUMN_SEQ = i - 1
                                Exit For
                            End If
                        Next
                    End If

                    .Cells("COLUMN_SEQ").Value = COLUMN_SEQ + 1
                    If COLUMN_SEQ <> 0 Then
                        Dim row As DataRow = dst.Tables("ASTDSQLS").Rows.Find(New Object() {Absx1.txtFor("FORM_NAME").Text, COLUMN_SEQ})
                        If .Cells("COLUMN_TYPE").Value & "" = "" Then
                            .Cells("COLUMN_TYPE").Value = row.Item("COLUMN_TYPE")
                        End If
                    End If

                    ' Sort_grdColumns(grdASTDSQLS, "COLUMN_SEQ")
                End If
            End If
        End With
    End Sub

    Private Sub grdASTDSQLV_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTDSQLV.BeforeRowUpdate
        With grdASTDSQLV.ActiveRow
            Select Case .Band.Key

                Case "ASTDSQLV"
                    If .Cells("FORM_NAME").Text = "" Then
                        .Cells("FORM_NAME").Value = Absx1.txtFor("FORM_NAME").Text
                        If Val(.Cells("VALUE_LIST_SEQ").Value & "") = 0 Then
                            Dim VALUE_LIST_SEQ As Integer = Val(dst.Tables("ASTDSQLV").Compute("MAX(VALUE_LIST_SEQ)", "") & "")
                            .Cells("VALUE_LIST_SEQ").Value = VALUE_LIST_SEQ + 1
                        End If
                    End If

                Case "ASTDSQLW"
                    'If .Cells("FORM_NAME").Text = "" Then
                    '.Cells("FORM_NAME").Value = Absx1.txtFor("FORM_NAME").Text
                    '.Cells("VALUE_LIST_NAME").Value = .ParentRow.Cells("VALUE_LIST_NAME").Text
                    If Val(.Cells("VALUE_LIST_CODE_SEQ").Value & "") = 0 Then
                        Dim VALUE_LIST_CODE_SEQ As Integer = Val(dst.Tables("ASTDSQLW").Compute("MAX(VALUE_LIST_CODE_SEQ)", "VALUE_LIST_NAME = '" & .ParentRow.Cells("VALUE_LIST_NAME").Text & "'") & "")
                        .Cells("VALUE_LIST_CODE_SEQ").Value = VALUE_LIST_CODE_SEQ + 1
                    End If
                    'End If
            End Select
        End With
    End Sub

    Private Sub cmdSchema_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSchema.Click

        If txtSCHEMA.Text.Length <> 3 Then
            MsgBox("No Schema Specified", MsgBoxStyle.OkOnly, "Cannot Copy Definition")
            Exit Sub
        End If

        Dim SCHEMA As String = txtSCHEMA.Text
        'If MsgBox("Use TST?", MsgBoxStyle.YesNo, "Answer Y to use TST, or N to use " & SCHEMA) = MsgBoxResult.Yes Then
        '    SCHEMA = "TST"
        'End If
        Dim LINK As String = txtLINK.Text
        If LINK <> "" Then LINK = "@" & LINK

        BeginTrans()
        Try
            For Each TABLE_NAME As String In New String() _
            {"ASTDSQLA", "ASTDSQLB", "ASTDSQLC", "ASTDSQLD", "ASTDSQLE", "ASTDSQLF",
             "ASTDSQLH", "ASTDSQLJ", "ASTDSQLS", "ASTDSQLV", "ASTDSQLW", "ASTDSQLX", "ASTDSQLY"}
                ASCMAIN1.sql = "Delete from " & SCHEMA & "." & TABLE_NAME & LINK _
                    & " where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Insert into " & SCHEMA & "." & TABLE_NAME & LINK _
                    & " Select * from " & TABLE_NAME _
                    & " where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
                ASCDATA1.ExecuteSQL()
            Next

            ASCMAIN1.sql = "Insert into " & SCHEMA & ".ASTDSQLK" & LINK _
            & " Select * from ASTDSQLK where COLUMN_NAME in " _
            & "(Select COLUMN_NAME from ASTDSQLA where FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "')" _
            & " and COLUMN_NAME Not in (Select COLUMN_NAME from " & SCHEMA & ".ASTDSQLK" & LINK & ")"
            ASCDATA1.ExecuteSQL()

            CommitTrans("Copy Successful")
        Catch ex As Exception
            Rollback("Error: " & ex.Message)
        End Try
    End Sub

    Overrides Sub Export_rows_to_SQL_custom()
        Export_rows_to_SQL("ASTDSQLA,ASTDSQLB,ASTDSQLC,ASTDSQLD,ASTDSQLE,ASTDSQLF,ASTDSQLH" _
                           & ",ASTDSQLJ,ASTDSQLS,ASTDSQLV,ASTDSQLW,ASTDSQLX,ASTDSQLY",
                           "FORM_NAME",
                           "",
                           "FORM_NAME", "ASTDSQLF")
    End Sub

    Private Sub grdASTDSQLS_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdASTDSQLS.AfterRowUpdate

    End Sub

    Private Sub grdASTDSQLS_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdASTDSQLS.AfterRowInsert

    End Sub
End Class