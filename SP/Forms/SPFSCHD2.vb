Public Class SPFSCHD2

    'Public APT As Infragistics.Win.UltraWinSchedule.Appointment
    Public rowSPTSCHD1 As DataRow
    Public UPDATED As Boolean = False
    Public SALES_DIVISION_CODE As String
    Public DEPT_CODE As String

    Private Sub SHFSCHD2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "Select SPTSCHD1.*, TATWHOP1.WH_OPER_NAME, TATWHOP1.WH_OPER_GRP, 0 TOTAL_DAYS from SPTSCHD1,TATWHOP1 where TATWHOP1.WH_OPER_ID = SPTSCHD1.WH_OPER_ID"
            Create_TDA(.Tables.Add, "SPTSCHD1", "**", 0)

            Create_Lookup("TATWHOP1", "*", "WH_OPER_ID = :PARM1 AND SALES_DIVISION_CODE = :PARM2", "VV", False)
            Create_Lookup("ASTUSER1", "*", "USER_ID = :PARM1 AND SALES_DIVISION_CODE = :PARM2", "VV", False)
            Create_Lookup("TATUSER1", "*", "USER_ID = :PARM1 AND SALES_DIVISION_CODE = :PARM2", "VV", False)

            Create_TDA(.Tables.Add, "SPTSCHDL", "*")
            Fill_Records("SPTSCHDL", String.Empty, True, "SELECT * FROM SPTSCHDL")

        End With

        Dim viewSPTSCHDL As DataView = New DataView(dst.Tables("SPTSCHDL"))
        cmbSCHED_CODE.ValueMember = "SCHED_CODE"
        cmbSCHED_CODE.DisplayMember = "SCHED_DESC"
        viewSPTSCHDL.Sort = "SCHED_DESC"
        cmbSCHED_CODE.DataSource = viewSPTSCHDL

        Dim SCHED_CODE As String = String.Empty
        Dim SCHED_DESC As String = String.Empty

        If rowSPTSCHD1 IsNot Nothing AndAlso (rowSPTSCHD1.Item("SCHED_CODE") & String.Empty).ToString.Trim <> String.Empty Then
            SCHED_CODE = (rowSPTSCHD1.Item("SCHED_CODE") & String.Empty).ToString.Trim
            If dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & SCHED_CODE & "'").Length > 0 Then
                viewSPTSCHDL.Find(SCHED_CODE)
                SCHED_DESC = dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & SCHED_CODE & "'")(0).Item("SCHED_DESC") & String.Empty
            End If
        End If

        If SCHED_DESC.Length > 0 Then
            cmbSCHED_CODE.Text = SCHED_DESC
        End If

        If rowSPTSCHD1 Is Nothing Then
            rowSPTSCHD1 = dst.Tables("SPTSCHD1").NewRow
            rowSPTSCHD1.Item("SCHED_DATE") = Now.Date
        End If

        If rowSPTSCHD1.Item("SCHED_DATE_END") & String.Empty = String.Empty Then
            rowSPTSCHD1.Item("SCHED_DATE_END") = rowSPTSCHD1.Item("SCHED_DATE")
        End If

        Dim row As DataRow = dst.Tables("SPTSCHD1").NewRow
        row.ItemArray = rowSPTSCHD1.ItemArray

        dst.Tables("SPTSCHD1").Rows.Add(row)

        If rowSPTSCHD1.Item("WH_OPER_NAME") & "" <> "" Then
            Me.Text = "Revise Scheduled Event for " & rowSPTSCHD1.Item("WH_OPER_NAME")
        End If

        Dim rowASTCODE1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ASTCODE1 WHERE TABLE_NAME = 'SPTSCHD1' AND COLUMN_NAME = 'DEPT_CODE' AND T_CODE = '" & DEPT_CODE & "'")

        If rowASTCODE1 IsNot Nothing Then
            lblPosition.Text = rowASTCODE1.Item("T_DESC") & String.Empty
        End If

        If DEPT_CODE = "OFF" Then
            MyBase.Absx1.SetABSViewName(txtWH_OPER_ID, "TAT_USER_ID")
            MyBase.Absx1.SetABSViewName(txtWH_OPER_NAME, "TAT_USER_ID")

            MyBase.Absx1.SetABSViewName(txtWH_OPER_GRP, "USER_OPER_GRP")
            MyBase.Absx1.SetABSViewName(txtWH_OPER_GRP_DESC, "USER_OPER_GRP")
        End If

        TABLE_NAME = "SPTSCHD1"
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim EMsg As String = ""

        If cmbSCHED_CODE.Text & "" = "" Then
            EMsg &= vbCr & "Reason Required"
        End If

        If Absx1.txtFor("WH_OPER_ID").Text = "" Then
            EMsg &= vbCr & "Whse Operator ID Required"
        Else
            Dim rowTATWHOP1 As DataRow = Nothing

            Dim WH_OPER_ID As String = Absx1.txtFor("WH_OPER_ID").Text.Trim
            Select Case DEPT_CODE
                Case "OFF"
                    rowTATWHOP1 = MyBase.LookUp("TATUSER1", New String() {WH_OPER_ID, SALES_DIVISION_CODE})
                Case Else
                    rowTATWHOP1 = MyBase.LookUp("TATWHOP1", New String() {WH_OPER_ID, SALES_DIVISION_CODE})
            End Select
            ' 
            If rowTATWHOP1 Is Nothing Then
                '     EMsg &= vbCr & "Invalid " & lblPosition.Text & " specified."
            End If
        End If

        If Not IsDate(dteSCHED_DATE.Value) OrElse Not IsDate(dteSCHED_DATE_END.Value) Then
            EMsg &= vbCr & "Start and End dates require valid dates."
        ElseIf dteSCHED_DATE.Value > dteSCHED_DATE_END.Value Then
            EMsg &= vbCr & "End date must be equal / greater Start date."
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Perform the Requested Action")
            Exit Sub
        End If

        dst.Tables("SPTSCHD1").Rows(0).Item("SCHED_CODE") = cmbSCHED_CODE.Value

        DATETIME_STAMP = DateTime.Now
        INIT_LAST("SPTSCHD1", True, , True)
        rowSPTSCHD1 = dst.Tables("SPTSCHD1").Rows(0)
        rowSPTSCHD1.Item("WH_OPER_GRP") = MyBase.Absx1.txtFor("WH_OPER_GRP").Text
        rowSPTSCHD1.Item("DEPT_CODE") = DEPT_CODE
        Update_Record_TDA("SPTSCHD1", "SCHED_NO = '" & rowSPTSCHD1.Item("SCHED_NO") & "'")
        UPDATED = True

        Me.Close()
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)
        Select Case COLUMN_NAME
            Case "WH_OPER_ID"
                If DEPT_CODE <> "WH" Then
                    sql_where = "TATUSER1.SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
                Else
                    sql_where = "SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
                End If
        End Select
    End Sub
End Class