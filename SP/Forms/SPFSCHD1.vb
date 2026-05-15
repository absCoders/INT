Public Class SPFSCHD1
    ' DRAG DROP NEEDS TO UPDATE DATATABLE AND DATABAES

    Dim SCHED_appts As New Dictionary(Of String, UltraWinSchedule.Appointment)
    Dim dvwSPTSCHD1 As DataView
    Dim deleted_rows() As String
    Dim apptEdit As Infragistics.Win.UltraWinSchedule.Appointment = Nothing
    Dim SALES_DIVISION_CODE As String
    Dim DEPT_CODE As String

    Dim whGroupValueList As New Infragistics.Win.ValueList
    Dim StatusValueList As New Infragistics.Win.ValueList

    Dim xWhGroupValueList As New Infragistics.Win.ValueList
    Dim xStatusValueList As New Infragistics.Win.ValueList

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SHFSCHDI" Then
            InquiryMode = True
            grdSPTSCHD1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        End If

        Dim SQL As String = String.Empty

        With dst
            ASCMAIN1.sql = "Select SPTSCHD1.*, TATWHOP1.WH_OPER_NAME, TATWHOP1.WH_OPER_GRP, 0 TOTAL_DAYS"
            ASCMAIN1.sql &= " from SPTSCHD1, TATWHOP1 "
            ASCMAIN1.sql &= " where TATWHOP1.WH_OPER_ID = SPTSCHD1.WH_OPER_ID"
            Create_TDA(.Tables.Add, "SPTSCHD1", "**", 0, True, String.Empty)
            .Tables("SPTSCHD1").Columns("WH_OPER_GRP").MaxLength = 10

            Create_TDA(.Tables.Add, "SPTSCHDL", "*")
            Fill_Records("SPTSCHDL", String.Empty, True, "SELECT * FROM SPTSCHDL")

            Create_TDA(.Tables.Add, "ASTCODE1", "*")
            Fill_Records("ASTCODE1", String.Empty, True, "SELECT * FROM ASTCODE1 WHERE TABLE_NAME = 'TATWHOP1' AND COLUMN_NAME = 'WH_OPER_GRP'")

        End With

        For Each rowASTCODE1 As DataRow In dst.Tables("ASTCODE1").Select("", "T_CODE")
            whGroupValueList.ValueListItems.Add(rowASTCODE1.Item("T_CODE").ToString.Trim, rowASTCODE1.Item("T_DESC").ToString.Trim)
            xWhGroupValueList.ValueListItems.Add(rowASTCODE1.Item("T_CODE").ToString.Trim, rowASTCODE1.Item("T_DESC").ToString.Trim)
        Next

        For Each rowSPTSCHDL As DataRow In dst.Tables("SPTSCHDL").Select("", "SCHED_CODE")
            StatusValueList.ValueListItems.Add(rowSPTSCHDL.Item("SCHED_CODE").ToString.Trim, rowSPTSCHDL.Item("SCHED_DESC").ToString.Trim)
            xStatusValueList.ValueListItems.Add(rowSPTSCHDL.Item("SCHED_CODE").ToString.Trim, rowSPTSCHDL.Item("SCHED_DESC").ToString.Trim)
        Next

        dvwSPTSCHD1 = New DataView(dst.Tables("SPTSCHD1"))
        grdSPTSCHD1.DataSource = dvwSPTSCHD1

        grdSPTSCHDX.DataSource = dst.Tables("SPTSCHD1")
        grdSPTSCHDL.DataSource = dst.Tables("SPTSCHDL")

        eventMonthView.VisibleWeeks = 4

        Dim calLook As Infragistics.Win.UltraWinSchedule.UltraCalendarLook = New Infragistics.Win.UltraWinSchedule.UltraCalendarLook
        calLook.AlternateMonthAppearance.BackColor = System.Drawing.Color.LightYellow
        calLook.SelectedDayAppearance.BackColor = System.Drawing.Color.DarkSlateBlue
        calLook.SelectedDayAppearance.ForeColor = System.Drawing.Color.YellowGreen

        Me.eventMonthView.CalendarLook = calLook

        dteStartDate.DateTime = DateAdd(DateInterval.Day, -30, DateTime.Now)
        dteEndDate.DateTime = DateAdd(DateInterval.Day, 30, DateTime.Now)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("SALES_DIVISION_CODE")
                Validate_Code("DEPT_CODE")

                If EMsg.Length > 0 Then
                    Exit Select
                End If

                If Not chkStartDate.Checked And Not IsDate(dteStartDate.DateTime) Then
                    EMsg &= vbCr & "You must provide a valid Start Date."
                ElseIf Not chkEnddate.Checked And Not IsDate(dteEndDate.DateTime) Then
                    EMsg &= vbCr & " You must provide a valid End Date."
                ElseIf Not chkStartDate.Checked AndAlso Not chkEnddate.Checked Then
                    If dteStartDate.DateTime.Date > dteEndDate.DateTime.Date Then
                        EMsg &= vbCr & "Start date must be less equal End date."
                    End If
                End If

                SALES_DIVISION_CODE = MyBase.Absx1.txtFor("SALES_DIVISION_CODE").Text
                DEPT_CODE = MyBase.Absx1.txtFor("DEPT_CODE").Text

                If EMsg.Length = 0 AndAlso Not InquiryMode Then
                    If Not ASCMAIN1.Logical_Lock("SPTSCHD1", SALES_DIVISION_CODE & "/" & DEPT_CODE) Then
                        Exit Sub
                    End If

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

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Status Legend").Visible = Not InquiryMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabSchedule.Visible = tf
        cmdRefresh.Enabled = tf
        grdSPTSCHDL.Enabled = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("SPTSCHD1").Rows.Clear()
        dst.EnforceConstraints = True

        SALES_DIVISION_CODE = String.Empty
        DEPT_CODE = String.Empty

        Absx1.txtFor("SALES_DIVISION_CODE").Clear()
        Absx1.txtFor("SALES_DIVISION_NAME").Clear()

        Absx1.txtFor("DEPT_CODE").Clear()
        Absx1.txtFor("DEPT_DESC").Clear()

        eventMonthView.CalendarInfo.Appointments.Clear()

    End Sub

    Sub Load_Record()
        Call ASCMAIN1.Progress("Now Loading Data")
        Call Save_Header_Fields(UltraGroupBox1)

        Me.GetAppointments()

        Me.eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        Me.eventMonthView.CalendarInfo.SelectedDateRanges.Add(DateTime.Now, 0)

        Setup_SPTSCHD1()

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdSPTSCHDX, "SS", "Show Filter", "Show GroupBox")
        'Load_Popup_Menu(eventMonthView, "B", "Edit Appointment")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If e.SourceControl.Name = "eventMonthView" Then
            Dim umv As UltraWinSchedule.UltraMonthViewSingle = DirectCast(e.SourceControl, Infragistics.Win.UltraWinSchedule.UltraMonthViewSingle)
            If umv.CalendarInfo.SelectedAppointments.Count = 1 Then
                apptEdit = umv.CalendarInfo.SelectedAppointments(0)
            Else
                apptEdit = Nothing
            End If
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.Key = "Edit Appointment" Then

            'Dim appt As UltraWinSchedule.Appointment = eventMonthView.GetAppointmentFromPoint(e.Cursor.Position)
            If apptEdit IsNot Nothing Then
                Dim SCHED_NO As String = apptEdit.Tag
                Edit_Appointment(SCHED_NO)
            End If
            Exit Sub
        End If

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

            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
            '    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

#End Region

#Region "Form Controls"

    Private Sub eventMonthView_BeforeAppointmentEdit(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinSchedule.BeforeAppointmentEditEventArgs) Handles eventMonthView.BeforeAppointmentEdit
        e.Cancel = True
        'Edit_Appointment(e.Appointment.Tag)
    End Sub

    Private Sub eventMonthView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles eventMonthView.Click
        Setup_SPTSCHD1()
    End Sub

    Private Sub eventMonthView_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles eventMonthView.MouseDoubleClick

        If InquiryMode Then Exit Sub
        If Not e.Button = MouseButtons.Left Then Exit Sub

        Dim point As System.Drawing.Point
        point = New System.Drawing.Point(e.X, e.Y)

        ' Determine where in the control the right button was pressed
        Dim objAppointment As Infragistics.Win.UltraWinSchedule.Appointment
        Dim objNote As Infragistics.Win.UltraWinSchedule.Note
        Dim objWeek As Infragistics.Win.UltraWinSchedule.Week
        Dim objDay As Infragistics.Win.UltraWinSchedule.Day
        Dim objDayOfWeek As Infragistics.Win.UltraWinSchedule.DayOfWeek

        ' See is we clicked an Appointment
        objAppointment = Me.eventMonthView.GetAppointmentFromPoint(e.X, e.Y)
        objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        If objAppointment Is Nothing AndAlso objDay Is Nothing Then
            Exit Sub
        End If

        If objDay IsNot Nothing AndAlso (objDay.Date < eventMonthView.CalendarInfo.MinDate OrElse objDay.Date > eventMonthView.CalendarInfo.MaxDate) Then
            MessageBox.Show("The selected date (" & objDay.Date & ") is out of specified Date Range.", "Create Appointment", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        objNote = Me.eventMonthView.GetNoteFromPoint(e.X, e.Y)
        objWeek = Me.eventMonthView.GetWeekFromPoint(e.X, e.Y)
        objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        objDayOfWeek = Me.eventMonthView.GetDayOfWeekFromPoint(e.X, e.Y)

        eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        eventMonthView.CalendarInfo.ActiveDay = objDay

        eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        eventMonthView.CalendarInfo.SelectedDateRanges.Add(objDay.Date, 0)
        Setup_SPTSCHD1()

        If objAppointment IsNot Nothing Then
            Dim SCHED_NO As String = objAppointment.Tag
            Edit_Appointment(SCHED_NO)
            Exit Sub
        End If

        'Me.eventMonthView.CalendarInfo.DA()
        Dim f As New SPFSCHD2
        f.rowSPTSCHD1 = dst.Tables("SPTSCHD1").NewRow
        f.rowSPTSCHD1.Item("SCHED_NO") = ASCMAIN1.Next_Control_No("SPTSCHD1.SCHED_NO")
        f.rowSPTSCHD1.Item("SCHED_DATE") = objDay.Date 'eventMonthView.CalendarInfo.ActiveDay.Date
        f.rowSPTSCHD1.Item("SCHED_DATE_END") = objDay.Date
        f.SALES_DIVISION_CODE = SALES_DIVISION_CODE
        f.DEPT_CODE = DEPT_CODE
        f.ShowDialog()

        If f.UPDATED Then
            Add_Appointment(f.rowSPTSCHD1)
            Dim SCHED_NO As String = f.rowSPTSCHD1.Item("SCHED_NO")
            Dim row As DataRow = dst.Tables("SPTSCHD1").Rows.Find(SCHED_NO)
            If row IsNot Nothing Then
                row.ItemArray = f.rowSPTSCHD1.ItemArray
            Else
                row = dst.Tables("SPTSCHD1").NewRow
                row.ItemArray = f.rowSPTSCHD1.ItemArray
                dst.Tables("SPTSCHD1").Rows.Add(row)
            End If
            dst.Tables("SPTSCHD1").AcceptChanges()
        End If

    End Sub


    Private Sub grdSPTSCHD1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTSCHD1.AfterRowsDeleted

        For i As Int16 = 0 To deleted_rows.Length - 1
            Dim SCHED_NO As String = deleted_rows(i)
            Remove_Appointment(SCHED_NO)
        Next
        Update_Record_TDA("SPTSCHD1")

    End Sub

    Private Sub grdSPTSCHD1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSPTSCHD1.AfterRowUpdate
        Update_Record_TDA("SPTSCHD1")
    End Sub

    Private Sub grdSPTSCHD1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTSCHD1.InitializeRow

        Dim rowSPTSCHDL As DataRow = dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & e.Row.Cells("SCHED_CODE").Value & "'")(0)

        Dim argbBackColor() As String = rowSPTSCHDL.Item("SCHED_BACKCOLOR").ToString.Split(".")
        Dim argbForeColor() As String = rowSPTSCHDL.Item("SCHED_FORECOLOR").ToString.Split(".")

        e.Row.Cells("SCHED_CODE").Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
        e.Row.Cells("SCHED_CODE").Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
    End Sub

    Private Sub grdSPTSCHD1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTSCHD1.DoubleClickRow
        Edit_Appointment(e.Row.Cells("SCHED_NO").Text)
    End Sub

    Private Sub grdSPTSCHD1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSPTSCHD1.BeforeRowsDeleted

        If InquiryMode Then
            e.Cancel = True
            Exit Sub
        End If

        Dim r As Int16 = grdSPTSCHD1.Selected.Rows.Count
        ReDim deleted_rows(r - 1)
        For i As Int16 = 0 To r - 1
            deleted_rows(i) = grdSPTSCHD1.Selected.Rows(i).Cells("SCHED_NO").Text
        Next
    End Sub

    Private Sub grdSPTSCHD1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTSCHD1.InitializeLayout

        grdSPTSCHD1.DisplayLayout.Bands(0).Columns("WH_OPER_GRP").Style = UltraWinGrid.ColumnStyle.DropDownList
        grdSPTSCHD1.DisplayLayout.Bands(0).Columns("WH_OPER_GRP").ValueList = whGroupValueList

        grdSPTSCHD1.DisplayLayout.Bands(0).Columns("SCHED_CODE").Style = UltraWinGrid.ColumnStyle.DropDownList
        grdSPTSCHD1.DisplayLayout.Bands(0).Columns("SCHED_CODE").ValueList = StatusValueList

    End Sub

    Private Sub grdSPTSCHD1_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdSPTSCHD1.AfterSelectChange
        Me.Setup_SPTSCHD1()
    End Sub


    Private Sub grdSPTSCHDX_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTSCHDX.InitializeLayout

        grdSPTSCHDX.DisplayLayout.Bands(0).Columns("WH_OPER_GRP").Style = UltraWinGrid.ColumnStyle.DropDownList
        grdSPTSCHDX.DisplayLayout.Bands(0).Columns("WH_OPER_GRP").ValueList = xWhGroupValueList

        grdSPTSCHDX.DisplayLayout.Bands(0).Columns("SCHED_CODE").Style = UltraWinGrid.ColumnStyle.DropDownList
        grdSPTSCHDX.DisplayLayout.Bands(0).Columns("SCHED_CODE").ValueList = xStatusValueList
    End Sub

    Private Sub grdSPTSCHDX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTSCHDX.InitializeRow
        Dim rowSPTSCHDL As DataRow = dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & e.Row.Cells("SCHED_CODE").Value & "'")(0)

        Dim argbBackColor() As String = rowSPTSCHDL.Item("SCHED_BACKCOLOR").ToString.Split(".")
        Dim argbForeColor() As String = rowSPTSCHDL.Item("SCHED_FORECOLOR").ToString.Split(".")

        e.Row.Cells("SCHED_CODE").Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
        e.Row.Cells("SCHED_CODE").Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
    End Sub

    Private Sub grdSPTSCHDL_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTSCHDL.AfterRowsDeleted
        Try
            MyBase.BeginTrans()
            Update_Record_TDA("SPTSCHDL")
            MyBase.CommitTrans()
        Catch ex As Exception
            MessageBox.Show("Error Deleting Statuses: " & ex.Message, "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.Rollback()
        End Try

    End Sub

    Private Sub grdSPTSCHDL_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSPTSCHDL.AfterRowUpdate

        Dim SCHED_CODE As String = e.Row.Cells("SCHED_CODE").Value
        Dim argbBackColor() As String = e.Row.Cells("SCHED_BACKCOLOR").Value.ToString.Split(".")
        Dim argbForeColor() As String = e.Row.Cells("SCHED_FORECOLOR").Value.ToString.Split(".")
        Dim SCHED_NO As String = String.Empty

        Dim rowAppts As DataRow() = dst.Tables("SPTSCHD1").Select("SCHED_CODE = '" & SCHED_CODE & "'")
        If rowAppts IsNot Nothing AndAlso rowAppts.Length > 0 Then
            Dim tbl As DataTable = dst.Tables("SPTSCHD1").Clone
            For Each row As DataRow In rowAppts
                tbl.ImportRow(row)
            Next

            For Each appt As Infragistics.Win.UltraWinSchedule.Appointment In eventMonthView.CalendarInfo.Appointments
                SCHED_NO = appt.Tag

                If tbl.Select("SCHED_NO = '" & SCHED_NO & "'").Length > 0 Then
                    appt.Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
                    appt.Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
                End If
            Next

        End If

        Dim rowSPTSCHDL As DataRow = dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & SCHED_CODE & "'")(0)
        If StatusValueList.ValueListItems.Contains(SCHED_CODE) Then
            StatusValueList.ValueListItems.Remove(rowSPTSCHDL.Item("SCHED_CODE"))
            xStatusValueList.ValueListItems.Remove(rowSPTSCHDL.Item("SCHED_CODE"))
        End If

        StatusValueList.ValueListItems.Add(SCHED_CODE, rowSPTSCHDL.Item("SCHED_DESC"))
        xStatusValueList.ValueListItems.Add(SCHED_CODE, rowSPTSCHDL.Item("SCHED_DESC"))

        Try
            MyBase.BeginTrans()
            Update_Record_TDA("SPTSCHDL")
            MyBase.CommitTrans()
        Catch ex As Exception
            MessageBox.Show("Error Updating Statuses: " & ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.Rollback()
        End Try

        grdSPTSCHD1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        grdSPTSCHDX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)


    End Sub

    Private Sub grdSPTSCHDL_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTSCHDL.BeforeRowUpdate

        If e.Row.Cells("BACKCOLOR").Value Is Nothing Then
            e.Row.Cells("BACKCOLOR").Value = System.Drawing.Color.White
        End If

        If e.Row.Cells("FORECOLOR").Value Is Nothing Then
            e.Row.Cells("FORECOLOR").Value = System.Drawing.Color.White
        End If

        Dim BACKCOLOR As System.Drawing.Color = e.Row.Cells("BACKCOLOR").Value
        Dim FORECOLOR As System.Drawing.Color = e.Row.Cells("FORECOLOR").Value
        Dim SCHED_DESC As String = (e.Row.Cells("SCHED_DESC").Value & String.Empty).ToString.Trim

        If BACKCOLOR = FORECOLOR Then
            e.Cancel = True
            MessageBox.Show("Backcolor and Forecolor must be different.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If SCHED_DESC.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Description is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim SCHED_BACKCOLOR As String = BACKCOLOR.A & "." & BACKCOLOR.R & "." & BACKCOLOR.G & "." & BACKCOLOR.B
        Dim SCHED_FORECOLOR As String = FORECOLOR.A & "." & FORECOLOR.R & "." & FORECOLOR.G & "." & FORECOLOR.B

        e.Row.Cells("SCHED_BACKCOLOR").Value = SCHED_BACKCOLOR
        e.Row.Cells("SCHED_FORECOLOR").Value = SCHED_FORECOLOR

        If e.Row.IsAddRow Then
            Dim codeFound As Boolean = False
            Dim SCHED_CODES As String = "ABCDEFGHIJLKMNOPQRSTUVWXYZ0123456789"
            For Each SCHED_CODE As Char In SCHED_CODES
                If dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & SCHED_CODE.ToString.Trim & "'").Length = 0 Then
                    e.Row.Cells("SCHED_CODE").Value = SCHED_CODE.ToString.Trim
                    codeFound = True
                    Exit For
                End If
            Next
            If Not codeFound Then
                e.Cancel = True
                MessageBox.Show("There are no codes available. Please contact ABS.", "Add Reason", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub grdSPTSCHDL_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTSCHDL.InitializeRow
        Dim argbBackColor() As String = e.Row.Cells("SCHED_BACKCOLOR").Value.ToString.Split(".")
        Dim argbForeColor() As String = e.Row.Cells("SCHED_FORECOLOR").Value.ToString.Split(".")

        e.Row.Cells("SCHED_BACKCOLOR").SelectedAppearance.ForeColor = System.Drawing.Color.Transparent
        e.Row.Cells("SCHED_FORECOLOR").SelectedAppearance.ForeColor = System.Drawing.Color.Transparent

        If e.Row.Cells("BACKCOLOR").Value IsNot Nothing Then
            Dim BACKCOLOR As System.Drawing.Color = e.Row.Cells("BACKCOLOR").Value
            Dim SCHED_BACKCOLOR As String = BACKCOLOR.A & "." & BACKCOLOR.R & "." & BACKCOLOR.G & "." & BACKCOLOR.B
            argbBackColor = SCHED_BACKCOLOR.Split(".")
        End If

        If e.Row.Cells("FORECOLOR").Value IsNot Nothing Then
            Dim FORECOLOR As System.Drawing.Color = e.Row.Cells("FORECOLOR").Value
            Dim SCHED_FORECOLOR As String = FORECOLOR.A & "." & FORECOLOR.R & "." & FORECOLOR.G & "." & FORECOLOR.B
            argbForeColor = SCHED_FORECOLOR.Split(".")
        End If

        If argbBackColor.Length <> 4 OrElse argbForeColor.Length <> 4 Then
            Exit Sub
        End If

        With e.Row.Cells("SCHED_DESC")
            .Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
            e.Row.Cells("BACKCOLOR").Value = .Appearance.BackColor
            .Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
            e.Row.Cells("FORECOLOR").Value = .Appearance.ForeColor
        End With
    End Sub

    Private Sub grdSPTSCHDL_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTSCHDL.InitializeLayout
        With grdSPTSCHDL.DisplayLayout.Bands(0).Columns("BACKCOLOR")
            .DataType = GetType(System.Drawing.Color)
            .Style = UltraWinGrid.ColumnStyle.Color
            .CellAppearance.ForeColor = System.Drawing.Color.Transparent
        End With

        With grdSPTSCHDL.DisplayLayout.Bands(0).Columns("FORECOLOR")
            .DataType = GetType(System.Drawing.Color)
            .Style = UltraWinGrid.ColumnStyle.Color
            .CellAppearance.ForeColor = System.Drawing.Color.Transparent
        End With

    End Sub

    Private Sub grdSPTSCHDL_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSPTSCHDL.BeforeRowsDeleted

        Dim SQL As String = "SELECT * FROM SPTSCHD1 WHERE SCHED_CODE = '" & grdSPTSCHDL.Selected.Rows(0).Cells("SCHED_CODE").Value & "'"
        If ASCDATA1.GetDataTable(SQL).Rows.Count > 0 Then
            e.Cancel = True
            MessageBox.Show("There are entries using this Status Code. Deletion denied!", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Private Sub grdSPTSCHDL_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSPTSCHDL.KeyPress
        Select Case grdSPTSCHDL.ActiveCell.Column.Key

            Case "BACKCOLOR", "FORECOLOR"
                e.KeyChar = vbTab
                e.Handled = True
        End Select
    End Sub


    Private Sub chkStartDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkStartDate.CheckedChanged
        dteStartDate.Enabled = Not chkStartDate.Checked
    End Sub

    Private Sub chkEnddate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEnddate.CheckedChanged
        dteEndDate.Enabled = Not chkEnddate.Checked
    End Sub


    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click

        Dim msg As String = String.Empty

        If Not chkStartDate.Checked And Not IsDate(dteStartDate.DateTime) Then
            msg &= vbCr & "You must provide a valid Start Date."
        ElseIf Not chkEnddate.Checked And Not IsDate(dteEndDate.DateTime) Then
            msg &= vbCr & " You must provide a valid End Date."
        ElseIf Not chkStartDate.Checked AndAlso Not chkEnddate.Checked Then
            If dteStartDate.DateTime.Date > dteEndDate.DateTime.Date Then
                msg &= vbCr & "Start date must be less equal End date."
            End If
        End If

        If msg.Length > 0 Then
            MessageBox.Show(msg, "Load Appointments", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Me.GetAppointments()
        End If

    End Sub

#End Region

#Region "Form Procedures"

    Sub Setup_SPTSCHD1()
        Dim SCHED_DATE As Date = eventMonthView.CalendarInfo.ActiveDay.Date
        dvwSPTSCHD1.RowFilter = "SCHED_DATE <= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "# AND SCHED_DATE_END >= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "#"
        grdSPTSCHD1.Text = "Scheduled Events for " & Format(SCHED_DATE, "MM/dd/yyyy")
    End Sub

    Sub Remove_Appointment(ByVal SCHED_NO As String)
        eventMonthView.CalendarInfo.Appointments.Remove(SCHED_appts(SCHED_NO))
        SCHED_appts.Remove(SCHED_NO)
    End Sub

    Sub Load_Appointments()
        For Each rowSPTSCHD1 As DataRow In dst.Tables("SPTSCHD1").Rows
            Add_Appointment(rowSPTSCHD1)
        Next
    End Sub

    Sub Add_Appointment(ByVal rowSPTSCHD1 As DataRow)
        Dim SUBJECT As String = rowSPTSCHD1.Item("WH_OPER_NAME") & String.Empty
        If rowSPTSCHD1.Item("SCHED_NOTE") & "" <> "" Then
            SUBJECT &= ":" & rowSPTSCHD1.Item("SCHED_NOTE") & String.Empty
        End If

        Dim SCHED_CODE As String = rowSPTSCHD1.Item("SCHED_CODE") & String.Empty
        Dim SCHED_NO As String = rowSPTSCHD1.Item("SCHED_NO") & String.Empty

        Dim appt As UltraWinSchedule.Appointment = Nothing 'New UltraWinSchedule.Appointment

        appt = eventMonthView.CalendarInfo.Appointments.Add(rowSPTSCHD1.Item("SCHED_DATE"), rowSPTSCHD1.Item("SCHED_DATE_END"), SUBJECT)
        appt.Tag = SCHED_NO

        Dim rowSPTSCHDL As DataRow = dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & rowSPTSCHD1.Item("SCHED_CODE") & "'")(0)

        Dim argbBackColor() As String = rowSPTSCHDL.Item("SCHED_BACKCOLOR").ToString.Split(".")
        Dim argbForeColor() As String = rowSPTSCHDL.Item("SCHED_FORECOLOR").ToString.Split(".")

        appt.Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
        appt.Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))

        appt.Description = SUBJECT
        appt.AllDayEvent = True

        If SCHED_appts.ContainsKey(SCHED_NO) Then
            SCHED_appts(SCHED_NO) = appt
        Else
            SCHED_appts.Add(SCHED_NO, appt)
        End If


    End Sub

    Sub Edit_Appointment(ByVal SCHED_NO As String)

        If SCHED_NO.Length = 0 Then Exit Sub

        Dim f As New SPFSCHD2
        f.rowSPTSCHD1 = dst.Tables("SPTSCHD1").Rows.Find(SCHED_NO)
        f.SALES_DIVISION_CODE = SALES_DIVISION_CODE
        f.DEPT_CODE = DEPT_CODE

        f.ShowDialog()

        If f.UPDATED Then
            Remove_Appointment(SCHED_NO)
            Add_Appointment(f.rowSPTSCHD1)
            Dim row As DataRow = dst.Tables("SPTSCHD1").Rows.Find(SCHED_NO)
            row.ItemArray = f.rowSPTSCHD1.ItemArray
            dst.Tables("SPTSCHD1").AcceptChanges()
        End If
    End Sub

    Sub GetAppointments()

        Dim sql As String = "Select * From SPTSCHD1"

        If DEPT_CODE = "WH" Then
            sql = "Select SPTSCHD1.*, TATWHOP1.WH_OPER_NAME, TATWHOP1.WH_OPER_GRP"
            sql &= ", ( "
            sql &= " NVL(SCHED_DATE_END, SCHED_DATE) - SCHED_DATE "
            sql &= " - "
            sql &= " ("
            sql &= " NEXT_DAY(NVL(SCHED_DATE_END, SCHED_DATE) - 7,'SUNDAY') - NEXT_DAY(SCHED_DATE - 1,'SUNDAY') "
            sql &= " + "
            sql &= " NEXT_DAY(NVL(SCHED_DATE_END, SCHED_DATE) - 7,'SATURDAY') - NEXT_DAY(SCHED_DATE - 1,'SATURDAY')"
            sql &= " ) /7 "
            sql &= " - 1"
            sql &= " ) TOTAL_DAYS "
            sql &= " from SPTSCHD1, TATWHOP1 "
            sql &= " where TATWHOP1.WH_OPER_ID = SPTSCHD1.WH_OPER_ID"
            sql &= " AND TATWHOP1.SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
        Else
            sql = "Select SPTSCHD1.*, ASTUSER1.USER_NAME WH_OPER_NAME, ASTUSER1.USER_COMPANY WH_OPER_GRP"
            sql &= ", ( "
            sql &= " NVL(SCHED_DATE_END, SCHED_DATE) - SCHED_DATE "
            sql &= " - "
            sql &= " ("
            sql &= " NEXT_DAY(NVL(SCHED_DATE_END, SCHED_DATE) - 7,'SUNDAY') - NEXT_DAY(SCHED_DATE - 1,'SUNDAY') "
            sql &= " + "
            sql &= " NEXT_DAY(NVL(SCHED_DATE_END, SCHED_DATE) - 7,'SATURDAY') - NEXT_DAY(SCHED_DATE - 1,'SATURDAY')"
            sql &= " ) /7 "
            sql &= " - 1"
            sql &= " ) TOTAL_DAYS "
            sql &= " from SPTSCHD1, ASTUSER1, TATUSER1 "
            sql &= " where ASTUSER1.USER_ID = SPTSCHD1.WH_OPER_ID"
            sql &= " AND ASTUSER1.USER_ID = TATUSER1.USER_ID"
            sql &= " AND TATUSER1.SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
        End If

        If Not chkStartDate.Checked Then
            sql &= " And SPTSCHD1.SCHED_DATE >= '" & dteStartDate.DateTime.ToString("dd-MMM-yyyy") & "'"
            eventMonthView.CalendarInfo.MinDate = dteStartDate.DateTime.Date
        Else
            eventMonthView.CalendarInfo.MinDate = CDate("01/01/09")
        End If

        If Not chkEnddate.Checked Then
            sql &= " And SPTSCHD1.SCHED_DATE <= '" & dteEndDate.DateTime.ToString("dd-MMM-yyyy") & "'"
            eventMonthView.CalendarInfo.MaxDate = dteEndDate.DateTime.Date
        Else
            eventMonthView.CalendarInfo.MaxDate = DateAdd(DateInterval.Year, 1, DateTime.Now)
        End If

        Fill_Records("SPTSCHD1", String.Empty, True, sql)
        Load_Appointments()


    End Sub

#End Region

End Class