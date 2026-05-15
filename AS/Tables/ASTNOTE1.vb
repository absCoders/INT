Public Class ASTNOTE1


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst
            Create_TDA(.Tables.Add, "ASTNOTE2", "*", 1)
            Create_TDA(.Tables.Add, "ASTNOTE3", "*", 1)
            .Tables("ASTNOTE3").Columns.Add("TABLE_COLUMN", GetType(System.String))

            Create_TDA(.Tables.Add, "ASTNOTE4", "*", 1)
        End With

        grdASTNOTE3.DataSource = dst.Tables("ASTNOTE3")
        grdASTNOTE4.DataSource = dst.Tables("ASTNOTE4")

        sql = "SELECT 'T' SEND_TYPE, 'To' SEND_DESC FROM DUAL"
        sql &= " UNION"
        sql &= " SELECT 'C' SEND_TYPE, 'CC' SEND_DESC FROM DUAL"
        sql &= " UNION"
        sql &= " SELECT 'B' SEND_TYPE, 'BCC' SEND_DESC FROM DUAL"
        sql &= " UNION"
        sql &= " SELECT 'F' SEND_TYPE, 'From' SEND_DESC FROM DUAL"

        ASCMAIN1.Add_Value_List(grdASTNOTE4, "SEND_TYPE", sql)

        Bind_Controls(Me, "ASTNOTE2")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey

            Case "Update"
                MyBase.Absx1.txtFor("NOTE_TEXT").Text = MyBase.Absx1.txtFor("NOTE_TEXT").Text.Trim
                If MyBase.Absx1.txtFor("NOTE_TEXT").Text.Length = 0 Then
                    EMsg &= vbCr & "The Email Text is required."
                End If

                MyBase.Absx1.txtFor("NOTE_DESC").Text = MyBase.Absx1.txtFor("NOTE_DESC").Text.Trim
                If MyBase.Absx1.txtFor("NOTE_DESC").Text.Length = 0 Then
                    EMsg &= vbCr & "The Note description is required."
                End If

                MyBase.Absx1.txtFor("EMAIL_SUBJECT").Text = MyBase.Absx1.txtFor("EMAIL_SUBJECT").Text.Trim
                If MyBase.Absx1.txtFor("EMAIL_SUBJECT").Text.Length = 0 Then
                    EMsg &= vbCr & "The Subject is required."
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim textNote As String = txtNOTE_TEXT.Text.Trim.ToUpper
        textNote &= " " & txtEMAIL_SUBJECT.Text.Trim.ToUpper

        ' Only keep the table/fields in teh document to be created
        For Each rowASTNOTE3 As DataRow In dst.Tables("ASTNOTE3").Select("", "")
            If Not textNote.Contains("{" & rowASTNOTE3.Item("TABLE_COLUMN") & "}") Then
                rowASTNOTE3.Delete()
            End If
        Next
        dst.Tables("ASTNOTE3").AcceptChanges()

        ' Reorder the SEND_LNO in the table so it is sequential
        Dim SEND_LNO As Int16 = 0
        For Each rowASTNOTE4 As DataRow In dst.Tables("ASTNOTE4").Select("", "SEND_LNO", DataViewRowState.CurrentRows)
            SEND_LNO += 1
            rowASTNOTE4.Item("SEND_LNO") = SEND_LNO
        Next
        dst.Tables("ASTNOTE4").AcceptChanges()

        Update_Record_TDA("ASTNOTE2", "NOTE_CODE = '" & MyBase.Absx1.txtFor("NOTE_CODE").Text & "'")
        Update_Record_TDA("ASTNOTE3", "NOTE_CODE = '" & MyBase.Absx1.txtFor("NOTE_CODE").Text & "'")
        Update_Record_TDA("ASTNOTE4", "NOTE_CODE = '" & MyBase.Absx1.txtFor("NOTE_CODE").Text & "'")

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        MyBase.EnforceConstraints(False)

        Dim sql As String = String.Empty

        Call Fill_Records("ASTNOTE2", New String() {Absx1.txtFor("NOTE_CODE").Text})
        Call Fill_Records("ASTNOTE3", New String() {Absx1.txtFor("NOTE_CODE").Text})
        Call Fill_Records("ASTNOTE4", New String() {Absx1.txtFor("NOTE_CODE").Text})

        ' Create blank row so the text for the note/email my be entered
        If dst.Tables("ASTNOTE2").Rows.Count = 0 Then
            dst.Tables("ASTNOTE2").Rows.Add(New Object() {Absx1.txtFor("NOTE_CODE").Text, ""})
        End If

        For Each rowASTNOTE3 As DataRow In dst.Tables("ASTNOTE3").Rows
            rowASTNOTE3.Item("TABLE_COLUMN") = rowASTNOTE3.Item("TABLE_NAME") & "." & rowASTNOTE3.Item("COLUMN_NAME")
        Next

        ' get all fields for the already used tables
        sql = "SELECT '" & Absx1.txtFor("NOTE_CODE").Text & "' NOTE_CODE"
        sql &= ", TABLE_NAME, COLUMN_NAME, NULL FIELD_FORMAT"
        sql &= ", TABLE_NAME || '.' || COLUMN_NAME TABLE_COLUMN"
        sql &= " FROM USER_TAB_COLUMNS"
        sql &= " WHERE TABLE_NAME IN (SELECT TABLE_NAME FROM ASTNOTE3 WHERE NOTE_CODE = '" & Absx1.txtFor("NOTE_CODE").Text & "')"
        sql &= " AND (TABLE_NAME, COLUMN_NAME) NOT IN (SELECT TABLE_NAME, COLUMN_NAME FROM ASTNOTE3 WHERE NOTE_CODE = '" & Absx1.txtFor("NOTE_CODE").Text & "')"
        Call Fill_Records("ASTNOTE3", String.Empty, False, sql)

        Sort_grdColumns(grdASTNOTE3, "TABLE_COLUMN")
        Sort_grdColumns(grdASTNOTE4, "SEND_LNO")

        MyBase.EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("ASTNOTE2").Rows.Clear()
            dst.Tables("ASTNOTE3").Rows.Clear()
            dst.Tables("ASTNOTE4").Rows.Clear()
            MyBase.EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        grdASTNOTE3.Enabled = tf
        splNote2.Enabled = tf

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        If Not tf Then
            tabNote.SelectedTab = tabNote.Tabs(0)
        End If

        tabNote.Enabled = tf
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdASTNOTE4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTNOTE4.BeforeRowUpdate
        Dim gMsg As String = String.Empty

        If e.Row.Cells("SEND_TYPE").Value = String.Empty Then
            gMsg &= vbCr & "The Type is required."
        End If

        e.Row.Cells("EMAIL_ADDRESS").Value = (e.Row.Cells("EMAIL_ADDRESS").Value & String.Empty).ToString.Trim

        'USE_SREP_CODE, ALT_SREP_CODE, CURRENT_USER, EMAIL_ADDRESS
        If Val(e.Row.Cells("USE_SREP_CODE").Value & String.Empty) = 0 _
            AndAlso Val(e.Row.Cells("ALT_SREP_CODE").Value & String.Empty) = 0 _
            AndAlso Val(e.Row.Cells("CURRENT_USER").Value & String.Empty) = 0 _
            AndAlso e.Row.Cells("EMAIL_ADDRESS").Text.Trim.Length = 0 Then
            gMsg &= vbCr & "Please provide an email address or check one of the email address options."

        End If

        If gMsg.Length > 0 Then
            MessageBox.Show(gMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("NOTE_CODE").Value = MyBase.Absx1.txtFor("NOTE_CODE").Text
            e.Row.Cells("SEND_LNO").Value = Val(dst.Tables("ASTNOTE4").Compute("MAX(SEND_LNO)", "") & String.Empty) + 1
        End If

    End Sub

    Private Sub cmdAddLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddLookup.Click

        Dim tableName As String = ASCMAIN1.Get_txt_from_User("Provide Table Name", "Get Lookup Table", False, 8)

        tableName = tableName.Trim.ToUpper
        If dst.Tables("ASTNOTE3").Select("TABLE_NAME = '" & tableName & "'").Length > 0 Then
            MessageBox.Show("Table (" & tableName & ") already exists in lookup.", "Lookup", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        ' get all fields for the already used tables
        Dim sql As String = String.Empty
        Sql = "SELECT '" & Absx1.txtFor("NOTE_CODE").Text & "' NOTE_CODE"
        Sql &= ", TABLE_NAME, COLUMN_NAME, NULL FIELD_FORMAT"
        Sql &= ", TABLE_NAME || '.' || COLUMN_NAME TABLE_COLUMN"
        Sql &= " FROM USER_TAB_COLUMNS"
        sql &= " WHERE TABLE_NAME = '" & tableName & "'"
        Call Fill_Records("ASTNOTE3", String.Empty, False, Sql)
        Sort_grdColumns(grdASTNOTE3, "TABLE_COLUMN")

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(txtNOTE_TEXT, "B", "Insert Field")
        Call Load_Popup_Menu(txtEMAIL_SUBJECT, "B", "Insert Field")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        Select Case e.Tool.Key

            Case "txtNOTE_TEXT", "txtEMAIL_SUBJECT"
                If grdASTNOTE3.Selected.Rows.Count = 0 Then
                    e.Cancel = True
                    Exit Sub
                End If
            Case Else
                'e.Cancel = True
                Exit Sub
        End Select

    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        'Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        Dim txt As UltraWinEditors.UltraTextEditor = Nothing

        If e.Tool.OwningMenu.Key.StartsWith("txt") Then
            Select Case e.Tool.OwningMenu.Key
                Case "txtNOTE_TEXT"
                    txt = MyBase.Absx1.CtlFor("NOTE_TEXT")
                Case "txtEMAIL_SUBJECT"
                    txt = MyBase.Absx1.CtlFor("EMAIL_SUBJECT")
                Case Else
                    Exit Sub
            End Select

        Else
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked


            Case Else
                If e.Tool.OwningMenu.Key.StartsWith("txt") Then
                    If grdASTNOTE3.Selected.Rows.Count = 0 Then
                        Exit Sub
                    End If

                    Dim valueToInsert As String = grdASTNOTE3.Selected.Rows(0).Cells("TABLE_COLUMN").Value & String.Empty
                    Dim SelectionStart As Int16 = txt.SelectionStart
                    valueToInsert = "{" & valueToInsert & "}"
                    txt.Text = txt.Text.Insert(SelectionStart, valueToInsert)
                    txt.SelectionStart = SelectionStart + valueToInsert.Length
                    Exit Sub
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

End Class

