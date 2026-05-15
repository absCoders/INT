Public Class SPFCOOPT

    'Public APT As Infragistics.Win.UltraWinSchedule.Appointment
    Public rowSPTCOOP6 As DataRow
    Public rowSPTCOOP1 As DataRow
    Public UPDATED As Boolean = False
    Public CUST_CODE As String
    Public VEHICLE_CODE As String
    Public TASK_ASSIGNED_TO As String
    Public AUTH_NO As String

    Public CUST_NAME As String
    Public VEHICLE_DESC As String
    Public BOOKING_NAME As String

    Public frmASFBASE0 As ASFBASE0
    Dim WORK_IDs As New Dictionary(Of String, String)

    Dim AUTH_TNO As Int32

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "SPTCOOP1", "*")
            Create_TDA(.Tables.Add, "SPTCOOP6", "*")
            With .Tables("SPTCOOP6").Columns
                ' .Add("STEP_ACTION_DATE_NAME")
                .Add("TASK_ACTION_DATE", GetType(System.DateTime))
                .Add("TASK_ACTION")
                ' .Add("TASK_COMPLETED")
            End With

            'Dim DT As New DataTable("POTWPDMA")
            'DT = frmASFBASE0.dst.Tables("SPTCOOP6").Clone

            Create_TDA(.Tables.Add, "SPTCOOP7", "*")
            ' Create_TDA(.Tables.Add, "SPTCOOP7", "*")
            '.Tables.Add(frmASFBASE0.dst.Tables("POTWPDM5"))
            '.Tables.Add(frmASFBASE0.dst.Tables("SPTCOOP6"))
            '.Tables.Add(frmASFBASE0.dst.Tables("SPTCOOP7"))
        End With

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("VEHICLE_CODE").Text = VEHICLE_CODE
        Absx1.txtFor("AUTH_NO").Text = AUTH_NO

        Absx1.txtFor("CUST_NAME").Text = CUST_NAME
        Absx1.txtFor("VEHICLE_DESC").Text = VEHICLE_DESC
        Absx1.txtFor("BOOKING_NAME").Text = BOOKING_NAME

        TABLE_NAME = "SPTCOOP6"

        grdASTATTA2.DataSource = frmASFBASE0.dst.Tables("ASTATTA2")

        Dim rowSPTCOOP6 As DataRow = dst.Tables("SPTCOOP6").NewRow
        rowSPTCOOP6.ItemArray = Me.rowSPTCOOP6.ItemArray

        AUTH_TNO = Val(rowSPTCOOP6.Item("AUTH_TNO") & "")

        Dim rowSPTCOOP1 As DataRow = frmASFBASE0.dst.Tables("SPTCOOP1").Rows.Find(New Object() {AUTH_NO})
        '      rowSPTCOOP1.Item("STEP_ACTION_DATE") = rowSPTCOOP1.Item("STEP_ACTION_DATE")

        Dim rowSPTCOOP1_new As DataRow = dst.Tables("SPTCOOP1").NewRow
        For Each DC As DataColumn In dst.Tables("SPTCOOP1").Columns
            rowSPTCOOP1_new.Item(DC.ColumnName) = rowSPTCOOP1.Item(DC.ColumnName)
        Next
        dst.Tables("SPTCOOP1").Rows.Add(rowSPTCOOP1_new)

        Dim sql As String = "AUTH_TNO = " & CStr(AUTH_TNO)

        'Dim WORK_ID_previous As String = ""
        'For Each row As DataRow In frmASFBASE0.dst.Tables("SPTCOOP7").Select(sql, "WORK_ID")
        '    Dim WORK_ID As String = row.Item("WORK_ID")
        '    WORK_IDs.Add(WORK_ID, WORK_ID_previous)
        '    WORK_ID_previous = WORK_ID
        'Next

        grdSPTCOOP7.DataSource = frmASFBASE0.dst.Tables("SPTCOOP7")

        Dim dvw As DataView

        dvw = DirectCast(grdSPTCOOP7.DataSource, DataTable).DefaultView
        dvw.RowFilter = sql
        Sort_grdColumns(grdSPTCOOP7, "WORK_ID")

        dvw = New DataView(frmASFBASE0.dst.Tables("SPTCOOP6"))
        dvw.RowFilter = "AUTH_TNO = " & rowSPTCOOP6.Item("AUTH_TNO")
        grdSPTCOOP7.DataSource = dvw
        grdSPTCOOP7.Text = "All Actions in Task " & CStr(AUTH_TNO) & ":" & rowSPTCOOP6.Item("TASK_DESC")
        Sort_grdColumns(grdSPTCOOP7, "INIT_DATE")

        lblTASK.Text = "Task " & rowSPTCOOP6.Item("AUTH_TNO")

        ASCMAIN1.sql = "" ' NECESSARY BECAUSE ASCMAIN1.SQL HAS STUFF IN IT 

        ASCMAIN1.Add_Value_List(grdSPTCOOP7, "TASK_STATUS")
        ' ASCMAIN1.Add_Value_List(grdSPTCOOP7, "TASK_STATUS")

        Absx1.txtFor("TASK_ACTION").Appearance.BackColor = Drawing.Color.LightGreen
        '   Absx1.chkFor("TASK_COMPLETED").Appearance.BackColor = Drawing.Color.LightGreen
        dteTASK_ACTION_DATE.Appearance.BackColor = Drawing.Color.LightGreen
        dteTASK_COMPLETED.Appearance.BackColor = Drawing.Color.LightGreen

        dteTASK_COMPLETED.Visible = False
        lblTASK_COMPLETED.Visible = False

        Me.Text = "Actions Performed on Task " & CStr(AUTH_TNO) & ": " & rowSPTCOOP6.Item("TASK_DESC") & " - " & rowSPTCOOP6.Item("TASK_DESC")

        Absx1.txtFor("TASK_ACTION").Focus()

        'If rowSPTCOOP6.Item("STEP_ACTION_DATE_NAME") & "" <> "" Then
        '    lblTASK_ACTION_DATE.Visible = True
        '    dteTASK_ACTION_DATE.Visible = True
        '    ' lblTASK_ACTION_DATE.Text = rowSPTCOOP6.Item("STEP_ACTION_DATE_NAME") & ""
        'End If

        lblNote.Visible = ASCMAIN1.USER_ID <> Absx1.txtFor("TASK_ASSIGNED_TO").Text

        If rowSPTCOOP6.Item("TASK_STATUS") & "" = "C" Then 'If optTASK_STATUS.Value = "C" Then
            cmdUpdate.Visible = False
            cmdCancel.Text = "Done"
            Set_Read_Only_for_ctl(optTASK_STATUS, True)
            Absx1.txtFor("TASK_ACTION").Visible = False
            chkComplete.Visible = False
            Set_Read_Only_for_ctl(Absx1.txtFor("TASK_NOTE"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("TASK_ASSIGNED_TO"), True)
            Set_Read_Only_for_ctl(Absx1.CtlFor("TASK_DUE"), True)
            Set_Read_Only_for_ctl(Absx1.CtlFor("STEP_ACTION_DATE"), True)
            lblTASK_ACTION.Visible = False
            '    lblTASK_ACTION .Visible = False
        End If

        If frmASFBASE0.EntryMode = "E" Or frmASFBASE0.EntryMode = "N" Then
        Else
            cmdUpdate.Visible = False
            cmdCancel.Text = "Done"
            Set_Read_Only(SplitContainer1.Panel1, True)
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim EMsg As String = ""
        If Absx1.txtFor("TASK_ASSIGNED_TO").Text = "" Then
            EMsg &= vbCr & "Task Owner is Required"
        Else
            If LookUp("ASTUSER1", Absx1.txtFor("TASK_ASSIGNED_TO").Text) Is Nothing Then
                EMsg &= vbCr & "Invalid Value specified for Task Owner"
            End If
        End If

        If Not IsDate(dteTASK_ASSIGNED.Value) OrElse Not IsDate(dteTASK_DUE.Value) Then
            EMsg &= vbCr & "Values for Date Assigned and Date Due are required"
        ElseIf Format(dteTASK_ASSIGNED.Value, "yyyyMMdd") > Format(dteTASK_DUE.Value, "yyyyMMdd") Then
            EMsg &= vbCr & "Date Due may not be prior to Date Assigned"
        End If

        Dim sql_prior_incomplete As String = "AUTH_TNO < " & CStr(AUTH_TNO) & " and ISNULL(TASK_STATUS,'?') <> 'C'"
        If chkComplete.Checked Then
            If frmASFBASE0.dst.Tables("SPTCOOP6").Select(sql_prior_incomplete).Length <> 0 Then
                If MsgBox("There are other Tasks scheduled before this Task that are not Complete Yet." _
                          & vbCrLf & "Do you want to Complete all of those prior Tasks in this single Update?", _
                          MsgBoxStyle.YesNo, "You are Completing a Task ahead of Prior Tasks") = MsgBoxResult.No Then
                    MsgBox("You Must Complete Tasks in Order", MsgBoxStyle.OkOnly, "Cannot Skip Ahead to Complete Tasks")
                    Exit Sub
                End If
            End If

            If rowSPTCOOP6.Item("TASK_ACTION_DATE") & "" <> "" Then
                If dteTASK_ACTION_DATE.Value & "" = "" Then
                    EMsg &= vbCr & "You Must Specify a Value for " & lblTASK_ACTION_DATE.Text
                End If
            End If
            If optTASK_STATUS.Value = "U" Then
                EMsg &= vbCr & "Cannot set Task Status to Unassigned if you are Completing the Task"
            End If
        Else
            If optTASK_STATUS.Value = "C" Then
                EMsg &= vbCr & "Cannot set Task Status to Complete using Option Set - use the Checkbox to indicate that the Task is Complete"
            End If
        End If


        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Perform the Requested Action")
            Exit Sub
        End If

        DATETIME_STAMP = DateTime.Now + ASCMAIN1.NowTSD
        Synch_TABLE_NAME("SPTCOOP6")
        INIT_LAST("SPTCOOP6", True, , True)

        Dim rowSPTCOOP6_work As DataRow = dst.Tables("SPTCOOP6").Rows(0)

        If rowSPTCOOP6_work.Item("TASK_STATUS") = "U" Then
            rowSPTCOOP6_work.Item("TASK_ASSIGNED_TO") = DBNull.Value
        End If

        Dim rowSPTCOOP6_base As DataRow = frmASFBASE0.dst.Tables("SPTCOOP6").Rows.Find _
                    (New Object() {rowSPTCOOP6_work.Item("STYLE_GROUP_NO"), _
                                   rowSPTCOOP6_work.Item("STEP_LNO"), _
                                   rowSPTCOOP6_work.Item("TASK_LNO")})

        rowSPTCOOP6_base.Item("LAST_DATE") = DATETIME_STAMP
        rowSPTCOOP6_base.Item("LAST_OPER") = ASCMAIN1.USER_ID

        rowSPTCOOP6_base.Item("TASK_NOTE") = rowSPTCOOP6_work.Item("TASK_NOTE")
        rowSPTCOOP6_base.Item("TASK_STATUS") = rowSPTCOOP6_work.Item("TASK_STATUS")
        rowSPTCOOP6_base.Item("TASK_DUE") = rowSPTCOOP6_work.Item("TASK_DUE")
        rowSPTCOOP6_base.Item("TASK_ASSIGNED_TO") = rowSPTCOOP6_work.Item("TASK_ASSIGNED_TO")

        If chkComplete.Checked Then
            rowSPTCOOP6_base.Item("TASK_STATUS") = "C"
            rowSPTCOOP6_base.Item("TASK_COMPLETED_BY") = ASCMAIN1.USER_ID
            rowSPTCOOP6_base.Item("TASK_COMPLETED") = rowSPTCOOP6_work.Item("TASK_COMPLETED")

            If frmASFBASE0.dst.Tables("SPTCOOP6").Select(sql_prior_incomplete).Length <> 0 Then
                For Each rowSPTCOOP6_prior As DataRow In frmASFBASE0.dst.Tables("SPTCOOP6").Select(sql_prior_incomplete)
                    rowSPTCOOP6_prior.Item("TASK_STATUS") = "C"
                    rowSPTCOOP6_prior.Item("TASK_COMPLETED_BY") = ASCMAIN1.USER_ID
                    rowSPTCOOP6_prior.Item("TASK_COMPLETED") = rowSPTCOOP6_work.Item("TASK_COMPLETED")

                    Dim rowSPTCOOP7_prior As DataRow = frmASFBASE0.dst.Tables("SPTCOOP7").NewRow
                    For Each dcol As DataColumn In frmASFBASE0.dst.Tables("SPTCOOP7").Columns
                        If frmASFBASE0.dst.Tables("SPTCOOP6").Columns.Contains(dcol.ColumnName) Then
                            rowSPTCOOP7_prior.Item(dcol.ColumnName) = rowSPTCOOP6_prior.Item(dcol.ColumnName)
                        End If
                    Next
                    ' rowSPTCOOP7_prior.Item("TASK_LNO") = rowSPTCOOP6_prior.Item("TASK_LNO")
                    rowSPTCOOP7_prior.Item("WORK_ID") = ASCMAIN1.Next_Control_No("SPTCOOP7.WORK_ID")
                    rowSPTCOOP7_prior.Item("TASK_ACTION") = "Completed by Task Id " & rowSPTCOOP6_base.Item("AUTH_TNO")
                    frmASFBASE0.dst.Tables("SPTCOOP7").Rows.Add(rowSPTCOOP7_prior)
                Next
            End If
        End If


        If rowSPTCOOP6_work.Item("STEP_ACTION_DATE_NAME") & "" <> "" AndAlso dteTASK_ACTION_DATE.Value & "" <> "" Then
            Dim rowPOTWPDM5_base As DataRow = frmASFBASE0.dst.Tables("POTWPDM5").Rows.Find _
                        (New Object() {rowSPTCOOP6_work.Item("STYLE_GROUP_NO"), _
                                       rowSPTCOOP6_work.Item("STEP_LNO")})
            rowPOTWPDM5_base.Item("STEP_ACTION_DATE") = dteTASK_ACTION_DATE.Value
        End If

        Dim rowSPTCOOP7_base As DataRow = frmASFBASE0.dst.Tables("SPTCOOP7").NewRow
        For Each dcol As DataColumn In frmASFBASE0.dst.Tables("SPTCOOP7").Columns
            If dst.Tables("SPTCOOP6").Columns.Contains(dcol.ColumnName) Then
                rowSPTCOOP7_base.Item(dcol.ColumnName) = rowSPTCOOP6_work.Item(dcol.ColumnName)
            End If
        Next
        rowSPTCOOP7_base.Item("WORK_ID") = ASCMAIN1.Next_Control_No("SPTCOOP7.WORK_ID")
        frmASFBASE0.dst.Tables("SPTCOOP7").Rows.Add(rowSPTCOOP7_base)

        'Update_Record_TDA("SPTCOOP6")
        'Update_Record_TDA("SPTCOOP7")
        'Update_Record_TDA("POTWPDM5")

        UPDATED = True

        Me.Close()
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)
        Select Case COLUMN_NAME
            'Case "WH_OPER_ID"
            '    If VEHICLE_CODE <> "WH" Then
            '        sql_where = "TATUSER1.SALES_DIVISION_CODE = '" & CUST_CODE & "'"
            '    Else
            '        sql_where = "SALES_DIVISION_CODE = '" & CUST_CODE & "'"
            '    End If
        End Select
    End Sub

    Private Sub grdSPTCOOP6_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOP7.InitializeRow
        If rowSPTCOOP6 Is Nothing Then Exit Sub
        If e.Row.Cells("AUTH_TNO").Value = rowSPTCOOP6.Item("AUTH_TNO") Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        End If
    End Sub

    Private Sub chkComplete_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkComplete.CheckedChanged
        dteTASK_COMPLETED.Visible = chkComplete.Checked
        lblTASK_COMPLETED.Visible = chkComplete.Checked

        If chkComplete.Checked Then
            dteTASK_COMPLETED.Value = Now.Date
        Else
            dteTASK_COMPLETED.Value = DBNull.Value
        End If
    End Sub

    Private Sub grdSPTCOOP7_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOP7.InitializeRow
        If WORK_IDs.Count = 0 Then Exit Sub
        Dim WORK_ID As String = e.Row.Cells("WORK_ID").Value & ""
        If WORK_IDs.ContainsKey(WORK_ID) Then
            Dim WORK_ID_previous As String = WORK_IDs(WORK_ID)
            If WORK_ID_previous <> "" Then
                Dim row As DataRow = frmASFBASE0.dst.Tables("SPTCOOP7").Select("WORK_ID = '" & WORK_ID_previous & "'")(0)
                For Each COLUMN_NAME As String In New String() {"TASK_DESC", "TASK_DIR", "TASK_NOTE", "TASK_STATUS", "TASK_ASSIGNED_TO", "TASK_ASSIGNED", "TASK_DUE", "TASK_COMPLETED_BY", "TASK_COMPLETED", "STEP_ACTION_DATE"}
                    If row.Item(COLUMN_NAME) & "" <> e.Row.Cells(COLUMN_NAME).Value & "" Then
                        e.Row.Cells(COLUMN_NAME).Appearance.BackColor = Drawing.Color.LightPink
                    End If
                Next
            End If
        End If
    End Sub
End Class