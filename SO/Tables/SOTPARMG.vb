Public Class SOTPARMG

    Dim sqlSOTFORMU As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTFORMU", "*", 0)
        End With

        grdSOTFORMU.DataSource = dst.Tables("SOTFORMU")

        Create_Summary(grdSOTFORMU, "ORDR_FORM_USER_NO", "Count")

        ASCMAIN1.Add_Value_List(grdSOTFORMU, "ORDR_FORM_USER_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive"})
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTFORMU, "SS", "Show Filter", "Show GroupBox")
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

        Select Case grd.Name

            'Case "grdSOTFORMU"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            'Case "Add Codes"
            'If grd.Name = "grdSOTFORMU" Then
            '    Add_Codes(grdSOTFORMU, "ICTITEM1", "ITEM_CODE", "Items")
            'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                'https://stackoverflow.com/questions/3028150/what-is-proper-regex-expression-for-swift-codes
                ' Dim rx As String = "[A-Z]{6,6}[A-Z2-9][A-NP-Z0-9]([A-Z0-9]{3,3}){0,1}" from JPMC spec
                ' Dim rx As String = "^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$"
                Dim rx_EMAIL As String = "\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z"
                Dim regex_EMAIL As New System.Text.RegularExpressions.Regex(rx_EMAIL, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                ' Regex regex = New Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");


                Dim rx_NAME As String = "^[a-zA-Z0-9_. ]+$" ' "/^[a-zA-Z0-9_.]+$/"
                Dim regex_NAME As New System.Text.RegularExpressions.Regex(rx_NAME, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                For Each rowSOTFORMU As DataRow In dst.Tables("SOTFORMU").Select("", "", DataViewRowState.Added + DataViewRowState.ModifiedCurrent)
                    Dim ORDR_FORM_USER_NO As String = rowSOTFORMU.Item("ORDR_FORM_USER_NO") & ""
                    Dim ORDR_FORM_USER_NAME As String = Trim(rowSOTFORMU.Item("ORDR_FORM_USER_NAME") & "")
                    Dim ORDR_FORM_USER_EMAIL As String = rowSOTFORMU.Item("ORDR_FORM_USER_EMAIL") & ""
                    Dim ORDR_FORM_USER_STATUS As String = rowSOTFORMU.Item("ORDR_FORM_USER_STATUS") & ""

                    If ORDR_FORM_USER_NAME = "" Then
                        EMsg &= vbCr & $"Missing Name for User {ORDR_FORM_USER_NO}"
                    Else
                        If Not regex_NAME.IsMatch(ORDR_FORM_USER_NAME) Then
                            EMsg &= vbCr & $"Invalid Name {ORDR_FORM_USER_NAME} for User {ORDR_FORM_USER_NO}"
                        Else
                            If dst.Tables("SOTFORMU").Select($"ORDR_FORM_USER_NAME = '{ORDR_FORM_USER_NAME}' AND ORDR_FORM_USER_STATUS = 'A'").Length > 1 Then
                                EMsg &= vbCr & $"Duplicate Name {ORDR_FORM_USER_NAME} for User {ORDR_FORM_USER_NO}"
                            End If

                            Dim activeNameRows = dst.Tables("SOTFORMU").Select($"ORDR_FORM_USER_NAME = '{ORDR_FORM_USER_NAME}' AND ORDR_FORM_USER_STATUS = 'A' AND ORDR_FORM_USER_NO <> '{ORDR_FORM_USER_NO}'")
                            If activeNameRows.Length > 0 AndAlso ORDR_FORM_USER_STATUS = "A" AndAlso rowSOTFORMU.RowState = DataRowState.Modified Then
                                EMsg &= vbCr & $"Cannot set user {ORDR_FORM_USER_NO} to Active. Another active user with the name {ORDR_FORM_USER_NAME} already exists."
                            End If
                        End If
                    End If

                    If ORDR_FORM_USER_EMAIL = "" Then
                        EMsg &= vbCr & $"Missing email address for User {ORDR_FORM_USER_NO}"
                    Else
                        If Not regex_EMAIL.IsMatch(ORDR_FORM_USER_EMAIL) Then
                            EMsg &= vbCr & $"Invalid email address {ORDR_FORM_USER_EMAIL} for User {ORDR_FORM_USER_NO}"
                        Else
                            If dst.Tables("SOTFORMU").Select($"ORDR_FORM_USER_EMAIL = '{ORDR_FORM_USER_EMAIL}' AND ORDR_FORM_USER_STATUS = 'A'").Length > 1 Then
                                EMsg &= vbCr & $"Duplicate email address {ORDR_FORM_USER_EMAIL} for User {ORDR_FORM_USER_NO}"
                            End If

                            Dim activeEmailRows = dst.Tables("SOTFORMU").Select($"ORDR_FORM_USER_EMAIL = '{ORDR_FORM_USER_EMAIL}' AND ORDR_FORM_USER_STATUS = 'A' AND ORDR_FORM_USER_NO <> '{ORDR_FORM_USER_NO}'")
                            If activeEmailRows.Length > 0 AndAlso ORDR_FORM_USER_STATUS = "A" AndAlso rowSOTFORMU.RowState = DataRowState.Modified Then
                                EMsg &= vbCr & $"Cannot set user {ORDR_FORM_USER_NO} to Active. Another active user with the email {ORDR_FORM_USER_EMAIL} already exists."
                            End If
                        End If
                    End If

                Next


        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim ORDR_FORM_CODE As String = Absx1.txtFor("ORDR_FORM_CODE").Text
        'Dim sqlDelete = "ORDR_FORM_CODE = '" & ORDR_FORM_CODE & "'"

        For Each rowSOTFORMU As DataRow In dst.Tables("SOTFORMU").Select("", "", DataViewRowState.Added + DataViewRowState.ModifiedCurrent)
            If rowSOTFORMU.RowState = DataRowState.Added Then
                rowSOTFORMU.Item("INIT_DATE") = DATETIME_STAMP
            End If
            rowSOTFORMU.Item("LAST_DATE") = DATETIME_STAMP
        Next
        Update_Record_TDA("SOTFORMU") ', sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("SOTFORMU") ', New String() {Absx1.txtFor("ORDR_FORM_CODE").Text})
        Sort_grdColumns(grdSOTFORMU, "ORDR_FORM_USER_NAME")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTFORMU"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSOTFORMU.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTFORMU}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdSOTFORMU"

    Private Sub grdSOTFORMU_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTFORMU.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "ITEM_CODE"
            '    grdCodeDesc(grdSOTFORMU, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
        End Select
    End Sub

    Private Sub grdSOTFORMU_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTFORMU.AfterRowActivate

        'With grdSOTFORMU.DisplayLayout.Bands(0).Columns("ITEM_CODE")
        '    If grdSOTFORMU.ActiveRow.IsAddRow Then
        '        .CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdSOTFORMU_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTFORMU.AfterRowsDeleted

    End Sub

    Private Sub grdSOTFORMU_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTFORMU.AfterRowUpdate

    End Sub

    Private Sub grdSOTFORMU_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTFORMU.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim ORDR_FORM_USER_NO As String = grow.Cells("ORDR_FORM_USER_NO").Value
            Dim rowSOTFORMU As DataRow = dst.Tables("SOTFORMU").Rows.Find(New String() {ORDR_FORM_USER_NO})
            If Not rowSOTFORMU.RowState = DataRowState.Added Then
                MsgBox("Cannot Delete Previously Updated Users - Instead, set Status to Inactive", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub grdSOTFORMU_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTFORMU.BeforeRowUpdate
        'Dim row As DataRow = LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_FORM_USER_NO").Value = ASCMAIN1.Next_Control_No("SOTFORMU.ORDR_FORM_USER_NO")
            e.Row.Cells("ORDR_FORM_USER_STATUS").Value = "A"

            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("INIT_DATE").Value = Now
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("LAST_DATE").Value = Now
        Else
            Dim ORDR_FORM_USER_NO As String = e.Row.Cells("ORDR_FORM_USER_NO").Value
            Dim rowSOTFORMU As DataRow = dst.Tables("SOTFORMU").Rows.Find(New String() {ORDR_FORM_USER_NO})
            If Not rowSOTFORMU.RowState = DataRowState.Added Then
                e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("LAST_DATE").Value = Now
            End If
        End If

    End Sub

    'Private Sub grdSOTFORMU_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTFORMU.ClickCellButton
    '    Select Case e.Cell.Column.Key
    '        Case "ITEM_CODE"
    '            Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "SOTFORMU", "ITEM_CODE")
    '            grdClickCellButton(grdSOTFORMU, sql_where, True)
    '    End Select
    'End Sub

#End Region

End Class