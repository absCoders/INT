Public Class SPFCOOP2

    'Public APT As Infragistics.Win.UltraWinSchedule.Appointment
    Public rowSPTCOOPX As DataRow
    Public UPDATED As Boolean = False
    Public CUST_CODE As String
    Public VEHICLE_CODE As String
    Public SREP_CODE As String
    Public BRAND_CODE As String
    Public sqlSPTCOOPX As String

    Private Sub SPFCOOP2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = sqlSPTCOOPX
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False)
        End With

        If rowSPTCOOPX Is Nothing Then
            rowSPTCOOPX = dst.Tables("SPTCOOPX").NewRow
            rowSPTCOOPX.Item("AUTH_DATE") = Now.Date
        End If

        Dim row As DataRow = dst.Tables("SPTCOOPX").NewRow
        row.ItemArray = rowSPTCOOPX.ItemArray
        row.Item("AUTH_NO") = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
        row.Item("AUTH_LNO") = 1
        dst.Tables("SPTCOOPX").Rows.Add(row)

        'If rowSPTCOOPX.Item("WH_OPER_NAME") & "" <> "" Then
        '    Me.Text = "Revise Scheduled Event for " & rowSPTCOOPX.Item("WH_OPER_NAME")
        'End If

        ' Dim rowASTCODE1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ASTCODE1 WHERE TABLE_NAME = 'SPTSCHD1' AND COLUMN_NAME = 'DEPT_CODE' AND T_CODE = '" & VEHICLE_CODE & "'")

        'If rowASTCODE1 IsNot Nothing Then
        '    lblPosition.Text = rowASTCODE1.Item("T_DESC") & String.Empty
        'End If

        'If VEHICLE_CODE = "OFF" Then
        '    MyBase.Absx1.SetABSViewName(txtCUST_CODE, "TAT_USER_ID")
        '    MyBase.Absx1.SetABSViewName(txtCUST_NAME, "TAT_USER_ID")

        '    MyBase.Absx1.SetABSViewName(txtVEHICLE_CODE, "USER_OPER_GRP")
        '    MyBase.Absx1.SetABSViewName(txtVEHICLE_DESC, "USER_OPER_GRP")
        'End If

        TABLE_NAME = "SPTCOOPX"
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim EMsg As String = ""
        If Absx1.txtFor("CUST_CODE").Text = "" Then
            EMsg &= vbCr & "Customer Code Required"
        Else
           
        End If

        If Not IsDate(dteDROP_DATE.Value) OrElse Not IsDate(dteIN_STORE_DATE.Value) Then
            EMsg &= vbCr & "Drop Date and In Store Date values are required"
        ElseIf dteDROP_DATE.Value > dteIN_STORE_DATE.Value Then
            EMsg &= vbCr & "In Store Date not be earlier than Drop Date"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Perform the Requested Action")
            Exit Sub
        End If

        '  dst.Tables("SPTSCHD1").Rows(0).Item("SCHED_CODE") = cmbSCHED_CODE.Value

        DATETIME_STAMP = DateTime.Now
        INIT_LAST("SPTSCHD1", True, , True)
        rowSPTCOOPX = dst.Tables("SPTSCHD1").Rows(0)
        rowSPTCOOPX.Item("WH_OPER_GRP") = MyBase.Absx1.txtFor("WH_OPER_GRP").Text
        rowSPTCOOPX.Item("DEPT_CODE") = VEHICLE_CODE
        Update_Record_TDA("SPTSCHD1", "SCHED_NO = '" & rowSPTCOOPX.Item("SCHED_NO") & "'")
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

    Private Sub UltraNumericEditor4_ValueChanged(sender As System.Object, e As System.EventArgs) Handles UltraNumericEditor4.ValueChanged

    End Sub
End Class