Public Class ASTALRT1


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst
            ASCMAIN1.sql = "SELECT * FROM ASTALRT2 WHERE ALERT_TYPE=:PARM1 OR :PARM1 IS NULL"
            Create_TDA(.Tables.Add, "ASTALRT2", "**",0,False,"V",1)

            ASCMAIN1.sql = "SELECT * FROM ASTALRT3 WHERE :PARM1=NVL(ALERT_TYPE,:PARM1) OR :PARM1 IS NULL"
            Create_TDA(.Tables.Add, "ASTALRT3", "**",0,True,"V",1)
        End With

        dst.Tables("ASTALRT3").Columns("RECIPIENT_ID").AllowDBNull = True
        dst.Tables("ASTALRT3").Columns("RECIPIENT_ID").AutoIncrement = True

        grdASTALRT2.DataSource = dst.Tables("ASTALRT2")
        grdASTALRT3.DataSource = dst.Tables("ASTALRT3")

        grdASTALRT3.DisplayLayout.Bands(0).Columns("ENABLED").DefaultCellValue = "1"

        sql = "SELECT 'WARN' ALERT_CLASS, 'Warning' CLASS_DESC FROM DUAL"
        sql &= " UNION"
        sql &= " SELECT 'CRIT' ALERT_CLASS, 'Critical' CLASS_DESC FROM DUAL"
        sql &= " UNION"
        sql &= " SELECT 'RSLV' ALERT_CLASS, 'Resolution' CLASS_DESC FROM DUAL"

        ASCMAIN1.Add_Value_List(grdASTALRT2, "ALERT_CLASS", sql)



        Fill_Records("ASTALRT2", {Absx1.txtFor("ALERT_TYPE").Text})
        Fill_Records("ASTALRT3", {Absx1.txtFor("ALERT_TYPE").Text})

        'Bind_Controls(Me, "ASTNOTE2")

    End Sub

#Region "Overrides"


    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey

            Case "New"
                EMsg &= "Contact ABS to add a new alert"


            Case "Update"
                'MyBase.Absx1.txtFor("NOTE_TEXT").Text = MyBase.Absx1.txtFor("NOTE_TEXT").Text.Trim
                
                If MyBase.Absx1.txtFor("ALERT_MESSAGE_TEMPLATE").Text.Length = 0 Then
                    EMsg &= vbCr & "Message template is required"
                End If

                If MyBase.Absx1.numFor("THRESHOLD_MIN_WARN").Value Is DBNull.Value AndAlso MyBase.Absx1.numFor("THRESHOLD_MAX_WARN").Value Is DBNull.Value Then
                    EMsg &= vbCr & "Min or max warning threshold value is required"
                End If

                If MyBase.Absx1.numFor("THRESHOLD_MIN_WARN").Value IsNot DBNull.Value AndAlso MyBase.Absx1.numFor("THRESHOLD_MIN_CRIT").Value IsNot DBNull.Value AndAlso
                    MyBase.Absx1.numFor("THRESHOLD_MIN_WARN").Value < MyBase.Absx1.numFor("THRESHOLD_MIN_CRIT").Value Then
                    EMsg &= vbCr & "Min critical threshold value must be <= min warning threshold value"
                End If

                If MyBase.Absx1.numFor("THRESHOLD_MAX_WARN").Value IsNot DBNull.Value AndAlso MyBase.Absx1.numFor("THRESHOLD_MAX_CRIT").Value IsNot DBNull.Value AndAlso
                    MyBase.Absx1.numFor("THRESHOLD_MAX_WARN").Value > MyBase.Absx1.numFor("THRESHOLD_MAX_CRIT").Value Then
                    EMsg &= vbCr & "Max warning threshold value must be <= max critical threshold value"
                End If

                If MyBase.Absx1.numFor("ALERT_RESEND_MINS_DELAY_WARN").Value IsNot DBNull.Value AndAlso MyBase.Absx1.numFor("ALERT_RESEND_MINS_DELAY_CRIT").Value IsNot DBNull.Value AndAlso
                    MyBase.Absx1.numFor("ALERT_RESEND_MINS_DELAY_WARN").Value < MyBase.Absx1.numFor("ALERT_RESEND_MINS_DELAY_CRIT").Value Then
                    EMsg &= vbCr & "Critical resend delay must be <= warning resend delay"
                End If

        End Select
    End Sub


    Overrides Sub Proceed_Update_Special_Pre()
        'Dim textNote As String = txtNOTE_TEXT.Text.Trim.ToUpper
        'textNote &= " " & txtEMAIL_SUBJECT.Text.Trim.ToUpper

        '' Only keep the table/fields in teh document to be created
        'For Each rowASTNOTE3 As DataRow In dst.Tables("ASTNOTE3").Select("", "")
        '    If Not textNote.Contains("{" & rowASTNOTE3.Item("TABLE_COLUMN") & "}") Then
        '        rowASTNOTE3.Delete()
        '    End If
        'Next
        'dst.Tables("ASTNOTE3").AcceptChanges()

        '' Reorder the SEND_LNO in the table so it is sequential
        'Dim SEND_LNO As Int16 = 0
        'For Each rowASTNOTE4 As DataRow In dst.Tables("ASTNOTE4").Select("", "SEND_LNO", DataViewRowState.CurrentRows)
        '    SEND_LNO += 1
        '    rowASTNOTE4.Item("SEND_LNO") = SEND_LNO
        'Next
        'dst.Tables("ASTNOTE4").AcceptChanges()

        Update_Record_TDA("ASTALRT3")
        'Update_Record_TDA("ASTNOTE3", "NOTE_CODE = '" & MyBase.Absx1.txtFor("NOTE_CODE").Text & "'")
        'Update_Record_TDA("ASTNOTE4", "NOTE_CODE = '" & MyBase.Absx1.txtFor("NOTE_CODE").Text & "'")

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        MyBase.EnforceConstraints(False)

        Dim sql As String = String.Empty
        Dim ALERT_TYPE As String = Absx1.txtFor("ALERT_TYPE").Text

        sql = $"SELECT '{ALERT_TYPE}' ALERT_TYPE, '{ALERT_TYPE}' ALERT_NAME FROM ASTALRT1 UNION SELECT '' ALERT_TYPE, 'All' FROM DUAL"
        ASCMAIN1.Add_Value_List(grdASTALRT3, "ALERT_TYPE", sql)

        Fill_Records("ASTALRT2", {ALERT_TYPE})
        Fill_Records("ASTALRT3", {ALERT_TYPE})

        Dim alertQuery As String = dst.Tables("ASTALRT1").Rows(0).Item("ALERT_QUERY")
        Dim dtASTALRTX As DataTable = ASCDATA1.GetDataTable(alertQuery)
        grdAlertResults.DataSource = dtASTALRTX

        
        '' Create blank row so the text for the note/email my be entered
        'If dst.Tables("ASTNOTE2").Rows.Count = 0 Then
        '    dst.Tables("ASTNOTE2").Rows.Add(New Object() {Absx1.txtFor("NOTE_CODE").Text, ""})
        'End If

        'For Each rowASTNOTE3 As DataRow In dst.Tables("ASTNOTE3").Rows
        '    rowASTNOTE3.Item("TABLE_COLUMN") = rowASTNOTE3.Item("TABLE_NAME") & "." & rowASTNOTE3.Item("COLUMN_NAME")
        'Next

        '' get all fields for the already used tables
        'sql = "SELECT '" & Absx1.txtFor("NOTE_CODE").Text & "' NOTE_CODE"
        'sql &= ", TABLE_NAME, COLUMN_NAME, NULL FIELD_FORMAT"
        'sql &= ", TABLE_NAME || '.' || COLUMN_NAME TABLE_COLUMN"
        'sql &= " FROM USER_TAB_COLUMNS"
        'sql &= " WHERE TABLE_NAME IN (SELECT TABLE_NAME FROM ASTNOTE3 WHERE NOTE_CODE = '" & Absx1.txtFor("NOTE_CODE").Text & "')"
        'sql &= " AND (TABLE_NAME, COLUMN_NAME) NOT IN (SELECT TABLE_NAME, COLUMN_NAME FROM ASTNOTE3 WHERE NOTE_CODE = '" & Absx1.txtFor("NOTE_CODE").Text & "')"
        'Call Fill_Records("ASTNOTE3", String.Empty, False, sql)

        'Sort_grdColumns(grdASTNOTE3, "TABLE_COLUMN")
        'Sort_grdColumns(grdASTALRT3, "SEND_LNO")

        MyBase.EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("ASTALRT2").Rows.Clear()
            dst.Tables("ASTALRT3").Rows.Clear()
            MyBase.EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        'grdASTNOTE3.Enabled = tf
        'splNote2.Enabled = tf

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        'If Not tf Then
        '    tabAlerts.SelectedTab = tabAlerts.Tabs(0)
        'End If

        'tabAlerts.Enabled = tf
        chkCurrentlyActive.Enabled = False

        grdAlertResults.Visible = tf
        lblAlertQueryResults.Visible = tf
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdASTALRT3_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTALRT3.AfterRowInsert
        'Dim gMsg As String = String.Empty
        e.Row.Cells("INIT_DATE").Value = Now
    End Sub

    Private Sub cmdAddLookup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 

        'Dim tableName As String = ASCMAIN1.Get_txt_from_User("Provide Table Name", "Get Lookup Table", False, 8)

        'tableName = tableName.Trim.ToUpper
        'If dst.Tables("ASTNOTE3").Select("TABLE_NAME = '" & tableName & "'").Length > 0 Then
        '    MessageBox.Show("Table (" & tableName & ") already exists in lookup.", "Lookup", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    Exit Sub
        'End If

        '' get all fields for the already used tables
        'Dim sql As String = String.Empty
        'Sql = "SELECT '" & Absx1.txtFor("NOTE_CODE").Text & "' NOTE_CODE"
        'Sql &= ", TABLE_NAME, COLUMN_NAME, NULL FIELD_FORMAT"
        'Sql &= ", TABLE_NAME || '.' || COLUMN_NAME TABLE_COLUMN"
        'Sql &= " FROM USER_TAB_COLUMNS"
        'sql &= " WHERE TABLE_NAME = '" & tableName & "'"
        'Call Fill_Records("ASTNOTE3", String.Empty, False, Sql)
        'Sort_grdColumns(grdASTNOTE3, "TABLE_COLUMN")

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(txtNOTE_TEXT, "B", "Insert Field")
        'Call Load_Popup_Menu(txtEMAIL_SUBJECT, "B", "Insert Field")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        Select Case e.Tool.Key

            Case "txtNOTE_TEXT", "txtEMAIL_SUBJECT"
                'If grdASTNOTE3.Selected.Rows.Count = 0 Then
                '    e.Cancel = True
                '    Exit Sub
                'End If
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
            'Select Case e.Tool.OwningMenu.Key
            '    Case "txtNOTE_TEXT"
            '        txt = MyBase.Absx1.CtlFor("NOTE_TEXT")
            '    Case "txtEMAIL_SUBJECT"
            '        txt = MyBase.Absx1.CtlFor("EMAIL_SUBJECT")
            '    Case Else
            '        Exit Sub
            'End Select

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

End Class

