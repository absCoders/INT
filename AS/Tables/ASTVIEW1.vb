Public Class ASTVIEW1

    Dim tblASTVIEWC As New DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ASTVIEW2", "*", 2)
            dst.Tables("ASTVIEW2").Columns.Add("COLUMN_NAME_RETURNED")
            dst.Tables("ASTVIEW2").Columns("COLUMN_NAME_RETURNED").ReadOnly = False

            Create_TDA(.Tables.Add, "ASTVIEW4", "*", 2)

            Create_TDA(.Tables.Add, "ASTVIEW5", "*", 2)
        End With
        grdASTVIEW2.DataSource = dst.Tables("ASTVIEW2")
        grdASTVIEW4.DataSource = dst.Tables("ASTVIEW4")
        grdASTVIEW5.DataSource = dst.Tables("ASTVIEW5")


        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select COLUMN_NAME, ORDINAL_POSITION COLUMN_ID" _
            & ", DATA_TYPE, CHARACTER_MAXIMUM_LENGTH DATA_LENGTH " _
            & " from INFORMATION_SCHEMA.COLUMNS" _
            & " WHERE TABLE_NAME = ''" _
            & " AND TABLE_CATALOG = '" & ASCMAIN1.DBS_COMPANY & "'" _
            & " order by ORDINAL_POSITION"
        Else
            ASCMAIN1.sql = "Select COLUMN_NAME, COLUMN_ID, DATA_TYPE, DATA_LENGTH from USER_TAB_COLUMNS where TABLE_NAME = '" & "" & "'"
        End If

        grdASTVIEWC.DataSource = ASCDATA1.GetDataTable
        grdASTVIEWC.DisplayLayout.Bands(0).SortedColumns.Add(grdASTVIEWC.DisplayLayout.Bands(0).Columns("COLUMN_ID"), False)

        With grdASTVIEW2.DisplayLayout.Bands(0)
            .Columns("COLUMN_POSITION").Header.Fixed = True
            .Columns("COLUMN_NAME").Header.Fixed = True
        End With


    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTVIEW2, "BB", "Set Return Value", "Set Order By")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Show Document"
                Dim PROP_DOC_FILENAME = grd.ActiveRow.Cells("PROP_DOC_FILENAME").Text
                Show_Document(ASCMAIN1.Folders("Archive") & "Proposals\Generated\" & PROP_DOC_FILENAME)


            Case "Set Return Value"
                Absx1.txtFor("COLUMN_NAME").Text = grd.ActiveRow.Cells("COLUMN_NAME").Text

            Case "Set Order By"
                Absx1.txtFor("ORDER_BY").Text = IIf(grd.ActiveRow.Cells("COLUMN_ALIAS").Text = "", grd.ActiveRow.Cells("COLUMN_NAME").Text, grd.ActiveRow.Cells("COLUMN_ALIAS").Text)


        End Select
    End Sub

#End Region

    Private Sub grdASTVIEW2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTVIEW2.AfterRowUpdate
        If e.Row.Cells("COLUMN_NAME_RETURNED").Text = "1" Then
            For Each row As DataRow In dst.Tables("ASTVIEW2").Rows
                If row.Item("COLUMN_NAME") <> e.Row.Cells("COLUMN_NAME").Text Then
                    row.Item("COLUMN_NAME_RETURNED") = ""
                End If
            Next
            Absx1.txtFor("COLUMN_NAME").Text = e.Row.Cells("COLUMN_NAME").Text
        Else
            If Absx1.txtFor("COLUMN_NAME").Text = e.Row.Cells("COLUMN_NAME").Text Then
                Absx1.txtFor("COLUMN_NAME").Text = ""
            End If
        End If
    End Sub

    Private Sub grdASTVIEW2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTVIEW2.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            For Each gc As UltraWinGrid.UltraGridColumn In .DisplayLayout.Bands(0).Columns
                If htbkey_COLUMN_NAMEs.Contains(gc.Key) Then
                    .ActiveRow.Cells(gc.Key).Value = DirectCast(htbkey_COLUMN_NAMEs(gc.Key), UltraWinEditors.UltraTextEditor).Text
                End If
            Next
            If .ActiveRow.Cells("COLUMN_POSITION").Text = "" Then
                .ActiveRow.Cells("COLUMN_POSITION").Value = .Rows.Count
            End If
        End With
    End Sub

    Private Sub grdASTVIEW4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTVIEW4.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            For Each gc As UltraWinGrid.UltraGridColumn In .DisplayLayout.Bands(0).Columns
                If htbkey_COLUMN_NAMEs.Contains(gc.Key) Then
                    .ActiveRow.Cells(gc.Key).Value = DirectCast(htbkey_COLUMN_NAMEs(gc.Key), UltraWinEditors.UltraTextEditor).Text
                End If
            Next
            If .ActiveRow.Cells("HIERARCHAL_VIEW").Text = "" Then
                .ActiveRow.Cells("HIERARCHAL_VIEW").Value = .Rows.Count
            End If
            If .ActiveRow.Cells("HIERARCHAL_VIEW_LEVEL").Text = "" Then
                .ActiveRow.Cells("HIERARCHAL_VIEW_LEVEL").Value = 1
            End If
        End With
    End Sub

    Private Sub grdASTVIEW5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTVIEW5.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            For Each gc As UltraWinGrid.UltraGridColumn In .DisplayLayout.Bands(0).Columns
                If htbkey_COLUMN_NAMEs.Contains(gc.Key) Then
                    .ActiveRow.Cells(gc.Key).Value = DirectCast(htbkey_COLUMN_NAMEs(gc.Key), UltraWinEditors.UltraTextEditor).Text
                End If
            Next
        End With
    End Sub

    Private Sub grdASTVIEWC_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdASTVIEWC.AfterSelectChange
        If EntryMode = "New" Or EntryMode = "Edit" Then
            If grdASTVIEWC.ActiveRow IsNot Nothing Then
                If grdASTVIEWC.ActiveRow.Selected Then
                    COLUMN_NAME = grdASTVIEWC.ActiveRow.Cells("COLUMN_NAME").Text
                    If dst.Tables("ASTVIEW2").Select("COLUMN_NAME = '" & COLUMN_NAME & "'").GetLength(0) = 0 Then
                        Call Add_View_Column(COLUMN_NAME)
                    End If
                End If
            End If
        End If
    End Sub

    Sub Add_View_Column(ByVal COLUMN_NAME As String, Optional ByVal COLUMN_CAPTION As String = "", Optional ByVal COLUMN_WIDTH As Integer = 100)
        If COLUMN_CAPTION = "" Then
            COLUMN_CAPTION = ASCMAIN1.Make_Caption(COLUMN_NAME)
        End If

        Dim row As DataRow = dst.Tables("ASTVIEW2").NewRow
        row.Item("VIEW_NAME") = Absx1.CtlFor("VIEW_NAME").Text
        row.Item("TABLE_NAME") = Absx1.CtlFor("TABLE_NAME").Text
        row.Item("COLUMN_POSITION") = Val(dst.Tables("ASTVIEW2").Compute("MAX(COLUMN_POSITION)", "") & "") + 1
        row.Item("COLUMN_NAME") = COLUMN_NAME
        row.Item("COLUMN_CAPTION") = COLUMN_CAPTION
        row.Item("COLUMN_WIDTH") = COLUMN_WIDTH
        dst.Tables("ASTVIEW2").Rows.Add(row)
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        dst.Tables("ASTVIEW2").AcceptChanges()
        Dim COLUMN_POSITION As Integer = 0
        For Each rowASTVIEW2 As DataRow In dst.Tables("ASTVIEW2").Select("", "COLUMN_POSITION")
            rowASTVIEW2.SetAdded()
            COLUMN_POSITION = COLUMN_POSITION + 1
            rowASTVIEW2.Item("COLUMN_POSITION") = COLUMN_POSITION
            If rowASTVIEW2.Item("COLUMN_NAME_RETURNED") & "" = "1" Or COLUMN_POSITION = 1 Then
                Absx1.txtFor("COLUMN_NAME").Text = rowASTVIEW2.Item("COLUMN_NAME")
                tblASFBASE1.Rows(0).Item("COLUMN_NAME") = rowASTVIEW2.Item("COLUMN_NAME")
            End If
        Next

        Dim sql_Delete As String = "VIEW_NAME = '" & Absx1.txtFor("VIEW_NAME").Text & "' and TABLE_NAME = '" & Absx1.txtFor("TABLE_NAME").Text & "'"
        Call Update_Record_TDA("ASTVIEW2", sql_Delete)
        Call Update_Record_TDA("ASTVIEW4", sql_Delete)
        Call Update_Record_TDA("ASTVIEW5", sql_Delete)
    End Sub

    Overrides Sub Show_Record_Special()
        Call Fill_Records("ASTVIEW2", New String() {Absx1.txtFor("VIEW_NAME").Text, Absx1.txtFor("TABLE_NAME").Text})

        If Absx1.txtFor("COLUMN_NAME").Text <> "" Then
            Dim row() As DataRow = dst.Tables("ASTVIEW2").Select("COLUMN_NAME = '" & Absx1.txtFor("COLUMN_NAME").Text & "'")
            If row.Length = 1 Then
                row(0).Item("COLUMN_NAME_RETURNED") = "1"
                row(0).AcceptChanges()
            End If
        End If


        Call Fill_Records("ASTVIEW4", New String() {Absx1.txtFor("VIEW_NAME").Text, Absx1.txtFor("TABLE_NAME").Text})
        Call Fill_Records("ASTVIEW5", New String() {Absx1.txtFor("VIEW_NAME").Text, Absx1.txtFor("TABLE_NAME").Text})

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select COLUMN_NAME, ORDINAL_POSITION COLUMN_ID" _
            & ", DATA_TYPE, CHARACTER_MAXIMUM_LENGTH DATA_LENGTH " _
            & " from INFORMATION_SCHEMA.COLUMNS" _
            & " WHERE TABLE_NAME = '" & Absx1.CtlFor("TABLE_NAME").Text & "'" _
            & " AND TABLE_CATALOG = '" & ASCMAIN1.DBS_COMPANY & "'"
        Else
            ASCMAIN1.sql = "Select COLUMN_NAME, COLUMN_ID, DATA_TYPE, DATA_LENGTH" _
            & " from USER_TAB_COLUMNS where TABLE_NAME = '" & Absx1.CtlFor("TABLE_NAME").Text & "'"
        End If

        tblASTVIEWC = ASCDATA1.GetDataTable
        grdASTVIEWC.DataSource = tblASTVIEWC
        grdASTVIEWC.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdASTVIEWC.DisplayLayout.Bands(0).SortedColumns.Add(grdASTVIEWC.DisplayLayout.Bands(0).Columns("COLUMN_ID"), False)

        If EntryMode = "New" Then
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = "Select COLUMN_NAME" _
                    & " from INFORMATION_SCHEMA.KEY_COLUMN_USAGE" _
                    & " where TABLE_NAME = '" & Absx1.CtlFor("TABLE_NAME").Text & "'" _
                    & " and TABLE_CATALOG = '" & ASCMAIN1.DBS_COMPANY & "'"
            Else
                ASCMAIN1.sql = "SELECT COLUMN_NAME FROM USER_IND_COLUMNS WHERE INDEX_NAME IN (" _
                    & " SELECT CONSTRAINT_NAME FROM USER_CONSTRAINTS " _
                    & " WHERE TABLE_NAME = '" & Absx1.CtlFor("TABLE_NAME").Text & "'" _
                    & "   AND CONSTRAINT_TYPE = 'P') " _
                    & " ORDER BY COLUMN_POSITION"
            End If

            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                COLUMN_NAME = row.Item("COLUMN_NAME")
                Call Add_View_Column(COLUMN_NAME, "Code")
                If COLUMN_NAME Like "*_CODE" Then
                    Mid$(COLUMN_NAME, COLUMN_NAME.Length - 4, 5) = "_DESC"
                    If tblASTVIEWC.Select("COLUMN_NAME = '" & COLUMN_NAME & "'").GetLength(0) <> 0 Then
                        Call Add_View_Column(COLUMN_NAME, "Description", 300)
                    End If
                    Mid$(COLUMN_NAME, COLUMN_NAME.Length - 4, 5) = "_NAME"
                    If tblASTVIEWC.Select("COLUMN_NAME = '" & COLUMN_NAME & "'").GetLength(0) <> 0 Then
                        Call Add_View_Column(COLUMN_NAME, "Name", 300)
                    End If
                End If
            Next
        End If

    End Sub

    Overrides Sub Clear_Record_Special()

        If SELECTION_NO = 0 Then Exit Sub
        If ScreenMode Then
            dst.Tables("ASTVIEW2").Rows.Clear()
            dst.Tables("ASTVIEW4").Rows.Clear()
            dst.Tables("ASTVIEW5").Rows.Clear()
        End If
        tblASTVIEWC.Clear()


    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTVIEW2.Enabled = tf
        grdASTVIEW4.Enabled = tf
        grdASTVIEW5.Enabled = tf
        grdASTVIEWC.Enabled = tf

        cmdSchema.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        txtSCHEMA.Visible = tf And EntryMode = "View" And ASCMAIN1.Running_in_VS
        Set_Read_Only_for_ctl(txtSCHEMA, Not txtSCHEMA.Visible)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        Call ASCMAIN1.Load_Views()
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        If Absx1.GetABSColumnName(txtctl) = "VIEW_NAME" Then
            Absx1.CtlFor("TABLE_NAME").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("TABLE_NAME")
        End If
    End Sub
#End Region

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click
        Absx1.txtFor("CODE_TABLE").Text = Absx1.txtFor("TABLE_NAME").Text
        Absx1.txtFor("CODE_COLUMN").Text = Absx1.txtFor("VIEW_NAME").Text
        Absx1.txtFor("COLUMN_NAME").Text = "T_CODE"

        If dst.Tables("ASTVIEW2").Rows.Count > 0 Then
            For I As Integer = dst.Tables("ASTVIEW2").Rows.Count - 1 To 0 Step -1
                dst.Tables("ASTVIEW2").Rows(I).Delete()
            Next
        End If

        Dim rowASTVIEW2 As DataRow

        rowASTVIEW2 = dst.Tables("ASTVIEW2").NewRow
        rowASTVIEW2.Item("VIEW_NAME") = Absx1.txtFor("VIEW_NAME").Text
        rowASTVIEW2.Item("TABLE_NAME") = Absx1.txtFor("TABLE_NAME").Text
        rowASTVIEW2.Item("COLUMN_POSITION") = 1
        rowASTVIEW2.Item("COLUMN_NAME") = "T_CODE"
        rowASTVIEW2.Item("COLUMN_ALIAS") = Absx1.txtFor("VIEW_NAME").Text
        rowASTVIEW2.Item("COLUMN_CAPTION") = "Code"
        rowASTVIEW2.Item("COLUMN_WIDTH") = 100
        rowASTVIEW2.Item("COLUMN_NAME_RETURNED") = "1"
        dst.Tables("ASTVIEW2").Rows.Add(rowASTVIEW2)

        rowASTVIEW2 = dst.Tables("ASTVIEW2").NewRow
        rowASTVIEW2.Item("VIEW_NAME") = Absx1.txtFor("VIEW_NAME").Text
        rowASTVIEW2.Item("TABLE_NAME") = Absx1.txtFor("TABLE_NAME").Text
        rowASTVIEW2.Item("COLUMN_POSITION") = 2
        rowASTVIEW2.Item("COLUMN_NAME") = "T_DESC"
        rowASTVIEW2.Item("COLUMN_ALIAS") = Absx1.txtFor("VIEW_NAME").Text & "_DESC"
        rowASTVIEW2.Item("COLUMN_CAPTION") = "Description"
        rowASTVIEW2.Item("COLUMN_WIDTH") = 300
        dst.Tables("ASTVIEW2").Rows.Add(rowASTVIEW2)

    End Sub

    Private Sub cmdSchema_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSchema.Click

        If txtSCHEMA.Text.Length <> 3 Then
            MsgBox("No Schema Specified", MsgBoxStyle.OkOnly, "Cannot Copy Definition")
            Exit Sub
        End If

        BeginTrans()
        Try
            For Each TABLE_NAME As String In New String() {"ASTVIEW1", "ASTVIEW2", "ASTVIEW4", "ASTVIEW5"}
                ASCMAIN1.sql = "Delete from " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & "@" & txtSCHEMA.Text _
                    & " where TABLE_NAME = '" & Absx1.txtFor("TABLE_NAME").Text & "'" _
                    & " and VIEW_NAME = '" & Absx1.txtFor("VIEW_NAME").Text & "'"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Insert into " & ASCMAIN1.DBS_COMPANY & "." & TABLE_NAME & "@" & txtSCHEMA.Text _
                    & " Select * FROM " & TABLE_NAME _
                    & " where TABLE_NAME = '" & Absx1.txtFor("TABLE_NAME").Text & "'" _
                    & " and VIEW_NAME = '" & Absx1.txtFor("VIEW_NAME").Text & "'"
                ASCDATA1.ExecuteSQL()
            Next
            CommitTrans("Copy Successful")
        Catch ex As Exception
            Rollback("Error: " & ex.Message)
        End Try

    End Sub

    Overrides Sub Export_rows_to_SQL_custom()
        Export_rows_to_SQL("ASTVIEW1,ASTVIEW2,ASTVIEW4,ASTVIEW5", _
                           "VIEW_NAME,TABLE_NAME", _
                           "", _
                           "VIEW_NAME", "ASTVIEW1")
    End Sub
End Class