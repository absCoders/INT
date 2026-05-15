Imports ABSolution

Public Class APF1099S

    Private _saved10991Query As String = String.Empty
    Private _payments10991Query As String = String.Empty

    Private _saved10992Query As String = String.Empty
    Private _payments10992Query As String = String.Empty

    Private _payments10991QueryVendCodes As String = String.Empty

    Private Const selectedOPS_YYYYPP As String = "RYP_YYYYPP"
    Private Const selectedMinAmount1099 As String = "_AMT_1099"

    Private _ops_YYYYPP As String = String.Empty
    Private _RYP As String = String.Empty
    Private _AMT_1099 As Double = 0

    Private Const ManualEntry = "Manual"
    Private eMode As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        _saved10991Query = "Select YEAR_ID, VEND_CODE, VEND_NAME, VEND_ADDR1, VEND_ADDR2, VEND_ADDR3" _
        & ", VEND_CITY, VEND_STATE, VEND_ZIP_CODE, VEND_COUNTRY, VEND_1099_BOX" _
        & ", VEND_TAX_ID, VEND_TAX_ID_TYPE, REPORTED_1099, TOTAL_PAID " _
        & " FROM APT10991 WHERE YEAR_ID = '" & selectedOPS_YYYYPP & "'"

        _payments10991Query = "SELECT '" & selectedOPS_YYYYPP & "' YEAR_ID, APTCHCK2.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_ADDR1, APTVEND1.VEND_ADDR2, APTVEND1.VEND_ADDR3" _
        & ", APTVEND1.VEND_CITY, APTVEND1.VEND_STATE, APTVEND1.VEND_ZIP_CODE, APTVEND1.VEND_COUNTRY, APTVEND1.VEND_1099_BOX" _
        & ", APTVEND1.VEND_TAX_ID, APTVEND1.VEND_TAX_ID_TYPE, SUM(NVL(APTINVH1.INV_1099_AMT, 0)) REPORTED_1099" _
        & ", SUM(NVL(APTCHCK2.INV_AMT_APPLIED, 0)) TOTAL_PAID " _
        & " FROM APTINVH1, APTVEND1, APTCHCK1, APTCHCK2" _
        & " WHERE APTINVH1.VEND_CODE = APTVEND1.VEND_CODE" _
        & " And (NVL(APTVEND1.VEND_TAX_ID, '*') <> '*' OR NVL(APTVEND1.VEND_1099_BOX, 0) <> 0 )" _
        & " AND NVL(APTVEND1.VEND_TAX_ID_TYPE , 'N') IN ('E', 'S')" _
        & " And APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
        & " And APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
        & " AND APTCHCK1.CHECK_STATUS <> 'V'" _
        & " AND APTCHCK1.OPS_YYYYPP LIKE '" & selectedOPS_YYYYPP & "%'" _
        & " GROUP BY '" & selectedOPS_YYYYPP & "', APTCHCK2.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_ADDR1, APTVEND1.VEND_ADDR2, APTVEND1.VEND_ADDR3" _
        & " , APTVEND1.VEND_CITY, APTVEND1.VEND_STATE, APTVEND1.VEND_ZIP_CODE, APTVEND1.VEND_COUNTRY, APTVEND1.VEND_1099_BOX" _
        & " , APTVEND1.VEND_TAX_ID, APTVEND1.VEND_TAX_ID_TYPE" _
        & " HAVING SUM(CASE WHEN APTCHCK2.INV_AMT_APPLIED < 0 THEN 0 ELSE APTCHCK2.INV_AMT_APPLIED END) >= " & selectedMinAmount1099 _
        & "  UNION" _
        & "   SELECT '" & selectedOPS_YYYYPP & "' YEAR_ID, APTVEND1.VEND_CODE, APTVEND1.VEND_NAME, APTVEND1.VEND_ADDR1, APTVEND1.VEND_ADDR2, APTVEND1.VEND_ADDR3" _
        & " , APTVEND1.VEND_CITY, APTVEND1.VEND_STATE, APTVEND1.VEND_ZIP_CODE, APTVEND1.VEND_COUNTRY, APTVEND1.VEND_1099_BOX" _
        & " , APTVEND1.VEND_TAX_ID, APTVEND1.VEND_TAX_ID_TYPE, 0 REPORTED_1099, 0 TOTAL_PAID " _
        & "  FROM APTVEND1" _
        & "  WHERE (NVL(APTVEND1.VEND_TAX_ID, '*') <> '*' OR NVL(APTVEND1.VEND_1099_BOX, 0) <> 0 )" _
        & "  AND NVL(APTVEND1.VEND_TAX_ID_TYPE , 'N') IN ('E', 'S')" _
        & "  AND VEND_CODE NOT IN (SELECT DISTINCT VEND_CODE FROM APTINVH1 WHERE OPS_YYYYPP LIKE '" & selectedOPS_YYYYPP & "%')"

        _saved10992Query = "Select YEAR_ID, VEND_CODE, CHECK_NUM, SEQ_NUM" _
        & ", INV_NUM, INV_DATE, VOUCHER_NO, AMT_PAID, INV_AMT, INV_1099_AMT" _
        & " FROM APT10992 WHERE YEAR_ID = '" & selectedOPS_YYYYPP & "'"

        _payments10992Query = "Select '" & selectedOPS_YYYYPP & "' YEAR_ID, APTINVH1.VEND_CODE, APTCHCK1.CHECK_NUM, APTCHCK2.SEQ_NUM" _
        & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.VOUCHER_NO, NVL(APTCHCK2.INV_AMT_APPLIED, 0) AMT_PAID" _
        & ", NVL(APTINVH1.INV_AMT, 0) INV_AMT, NVL(APTINVH1.INV_1099_AMT, 0) INV_1099_AMT" _
        & " FROM APTINVH1, APTCHCK2, APTCHCK1, APTVEND1" _
        & " WHERE APTINVH1.VEND_CODE = APTVEND1.VEND_CODE" _
        & " And APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
        & " And APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
        & " AND APTCHCK1.CHECK_STATUS <> 'V'" _
        & " AND APTCHCK1.OPS_YYYYPP LIKE '" & selectedOPS_YYYYPP & "%'" _
        & " AND APTINVH1.VEND_CODE IN (SELECT VEND_CODE FROM _payments10991QueryVendCodes)"

        Dim sql As String = String.Empty

        With dst

            MyBase.Create_TDA(.Tables.Add, "APT10991", "*")
            .Tables("APT10991").Columns.Add("VALID_TAX_ID", GetType(System.Boolean), "VEND_1099_BOX IS NULL OR VEND_TAX_ID IS NULL OR VEND_TAX_ID_TYPE IS NULL")

            MyBase.Create_TDA(.Tables.Add, "APT10991rpt", "Select * From APT10991", -1, False)
            .Tables("APT10991rpt").Columns.Add("VALID_TAX_ID", GetType(System.Boolean), "VEND_1099_BOX IS NULL OR VEND_TAX_ID IS NULL OR VEND_TAX_ID_TYPE IS NULL")

            MyBase.Create_TDA(.Tables.Add, "APT10992", "*")

            .Relations.Add("APT10991_APT10992", _
            New DataColumn() {.Tables("APT10991").Columns("YEAR_ID"), .Tables("APT10991").Columns("VEND_CODE")}, _
            New DataColumn() {.Tables("APT10992").Columns("YEAR_ID"), .Tables("APT10992").Columns("VEND_CODE")})

            .Tables.Add("GLTYEARX")
            .Tables("GLTYEARX").Columns.Add("YEAR", GetType(System.String))
            For year As Long = Val(ASCMAIN1.CYP.Substring(0, 4)) To Val(ASCMAIN1.CYP.Substring(0, 4)) - 10 Step -1
                Dim row As DataRow = .Tables("GLTYEARX").NewRow
                row.Item("YEAR") = year.ToString
                .Tables("GLTYEARX").Rows.Add(row)
            Next

            Me.grdAPT10991.DataSource = dst.Tables("APT10991")
            Me.grdAPT10992.DataSource = dst.Tables("APT10992")
            Me.cmbYear.DataSource = dst.Tables("GLTYEARX")

            With Me.grdAPT10991.DisplayLayout.Bands(0).SortedColumns
                .Clear()
                .Add("VEND_CODE", False)
            End With

            With Me.grdAPT10992.DisplayLayout.Bands(0).SortedColumns
                .Clear()
                .Add("VEND_CODE", False)
                .Add("CHECK_NUM", False)
                .Add("SEQ_NUM", False)
                .Add("INV_NUM", False)
            End With

            Me.cmbYear.SelectedRow = Me.cmbYear.Rows(0)

            MyBase.Create_TDA(.Tables.Add, "APTPARM1", "*")
            MyBase.Fill_Records("APTPARM1", String.Empty, True, "SELECT * FROM APTPARM1 WHERE AP_PARM_KEY = 'Z'")

            MyBase.Create_TDA(.Tables.Add, "APTINVH1", "*")

        End With

        Me.Mode_Settings(False)

        Call Create_Summary(grdAPT10991, "VEND_CODE", "Count")
        Call Create_Summary(grdAPT10991, "TOTAL_PAID", "Sum")
        Call Create_Summary(grdAPT10991, "REPORTED_1099", "Sum")

        Call Create_Summary(grdAPT10992, "CHECK_NUM", "Count")
        Call Create_Summary(grdAPT10992, "AMT_PAID", "Sum")
        Call Create_Summary(grdAPT10992, "INV_AMT", "Sum")
        Call Create_Summary(grdAPT10992, "INV_1099_AMT", "Sum")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                eMode = "L"
                Me._AMT_1099 = Absx1.numFor("AMT_1099").Value
                If Me._AMT_1099 < 0 Then
                    EMsg = "Minimum amount to print must be greater equal 0."
                End If

            Case "Regenerate"
                Me._AMT_1099 = Absx1.numFor("AMT_1099").Value
                If Me._AMT_1099 < 0 Then
                    EMsg = "Minimum amount to print must be greater equal 0."
                End If
                eMode = "R"

            Case "Cancel"

            Case "Update"

            Case "Print 1099s"
                If dst.Tables("APT10991").Select("", "", DataViewRowState.ModifiedOriginal).Length > 0 Then
                    EMsg = "Changes must be saved before you can print 1099s."
                ElseIf dst.Tables("APT10992").Select("", "", DataViewRowState.Added).Length > 0 Then
                    EMsg = "Changes must be saved before you can print 1099s."
                ElseIf dst.Tables("APT10992").Select("", "", DataViewRowState.ModifiedOriginal).Length > 0 Then
                    EMsg = "Changes must be saved before you can print 1099s."
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                MyBase.EntryMode = "E"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Regenerate"
                MyBase.EntryMode = "R"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)

            Case "Print 1099s"
                Me.SetupPrintRecords()
                MyBase.Print_Report_Begin()
                Generate_Report("APR1099S")
                MyBase.Print_Report_End()

            Case "Add Payment"
                Me.AddPayment()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Regenerate").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Print 1099s").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Add Payment").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            If dst.Tables("APT10991").Rows.Count = 0 Then
                UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = not_iScreenMode
                UltraExplorerBar1.Groups("Screen Control").Items("Print 1099s").Settings.Enabled = not_iScreenMode
                UltraExplorerBar1.Groups("Screen Control").Items("Add Payment").Settings.Enabled = not_iScreenMode
            ElseIf dst.Tables("APT10991").Select("VALID_TAX_ID <> 0").Length > 0 Then
                UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = not_iScreenMode
            End If

            If eMode = "R" Then
                UltraExplorerBar1.Groups("Screen Control").Items("Print 1099s").Settings.Enabled = not_iScreenMode
            End If
        Else
            Me.Clear_Record()
        End If

        Me.SplitContainer1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("APT10991").Rows.Clear()
        dst.Tables("APT10992").Rows.Clear()
        dst.Tables("APTINVH1").Rows.Clear()
        dst.Tables("APT10991rpt").Rows.Clear()
        dst.EnforceConstraints = True
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading 1099 information")

        dst.EnforceConstraints = False

        Me._RYP = MyBase.Absx1.cmbFor("OPS_YYYYPP").Text
        Me._AMT_1099 = Absx1.numFor("AMT_1099").Value

        Dim sql10991 As String = String.Empty
        Dim sql10992 As String = String.Empty
        Dim periodField As String = selectedOPS_YYYYPP
        Dim sql As String = String.Empty

        Select Case MyBase.EntryMode

            Case "E"
                sql10991 = _saved10991Query
                sql10992 = _saved10992Query

                sql10991 = sql10991.Replace(periodField, _RYP)
                sql10992 = sql10992.Replace(periodField, _RYP)

            Case "R"
                sql10991 = _payments10991Query

                sql10991 = sql10991.Replace(periodField, _RYP)
                sql10991 = sql10991.Replace(selectedMinAmount1099, Me._AMT_1099)

                If _payments10991QueryVendCodes.Length = 0 Then
                    _payments10991QueryVendCodes = ASCMAIN1.Temp_Table(sql10991)
                Else
                    sql = "TRUNCATE TABLE " & _payments10991QueryVendCodes
                    ASCDATA1.ExecuteSQL(sql)
                    sql = "INSERT INTO " & _payments10991QueryVendCodes & " " & sql10991
                    ASCDATA1.ExecuteSQL(sql)
                End If

                sql10991 = "SELECT * FROM " & _payments10991QueryVendCodes

                sql10992 = _payments10992Query
                sql10992 = sql10992.Replace(periodField, _RYP)
                sql10992 = sql10992.Replace("_payments10991QueryVendCodes", _payments10991QueryVendCodes)

        End Select

        MyBase.Fill_Records("APT10991", String.Empty, True, sql10991)
        MyBase.Fill_Records("APT10992", String.Empty, True, sql10992)

        If Me.grdAPT10991.Rows.Count > 0 Then
            Me.grdAPT10991.ActiveRow = Me.grdAPT10991.Rows(0)
        End If

        If Me.grdAPT10992.Rows.Count > 0 Then
            Dim _APTINVH1sql As String = String.Empty
            For Each dr As DataRow In dst.Tables("APT10992").Rows
                _APTINVH1sql &= ", '" & dr.Item("VOUCHER_NO") & "' "
            Next
            _APTINVH1sql = _APTINVH1sql.Substring(1)
            _APTINVH1sql = "Select * From APTINVH1 WHERE VOUCHER_NO IN (" & _APTINVH1sql & ")"
            MyBase.Fill_Records("APTINVH1", String.Empty, True, _APTINVH1sql)
        End If

        dst.EnforceConstraints = True

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()
        MyBase.BeginTrans()

        Dim sql As String = String.Empty

        sql = "Delete From APT10991 WHERE YEAR_ID = '" & _RYP & "'"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Delete From APT10992 WHERE YEAR_ID = '" & _RYP & "'"
        ASCDATA1.ExecuteSQL(sql)

        For Each row As DataRow In dst.Tables("APT10991").Rows
            row.AcceptChanges()
            row.SetAdded()
        Next

        For Each row As DataRow In dst.Tables("APT10992").Rows
            row.AcceptChanges()
            row.SetAdded()
        Next

        MyBase.Update_Record_TDA("APT10991")
        MyBase.Update_Record_TDA("APT10992")
        MyBase.Update_Record_TDA("APTINVH1")

        MyBase.CommitTrans("Update Complete")
    End Sub

    Private Sub AddPayment()

        Dim VEND_CODE As String = Me.grdAPT10991.ActiveRow.Cells("VEND_CODE").Value
        Dim VEND_NAME As String = Me.grdAPT10991.ActiveRow.Cells("VEND_NAME").Value
        Dim SEQ_NUM As String = "0"

        Dim dd As Double = 0
        Dim userValue As String = InputBox$("Enter payment value for vendor (" & VEND_CODE & ") " & VEND_NAME & ".", "Payment", "0.00")
        Double.TryParse(userValue, dd)

        If dd = 0 Then Exit Sub

        Dim dr As DataRow = dst.Tables("APT10992").NewRow

        Try
            Dim sql As String = "YEAR_ID = '" & _RYP & "' AND VEND_CODE = '" & VEND_CODE & "' AND CHECK_NUM = '" & ManualEntry & "'"
            For Each rowAPT10992 As DataRow In dst.Tables("APT10992").Select(sql, "SEQ_NUM")
                If Val(rowAPT10992.Item("SEQ_NUM") & String.Empty) > Val(SEQ_NUM) Then
                    SEQ_NUM = rowAPT10992.Item("SEQ_NUM") & String.Empty
                End If
            Next
        Catch ex As Exception

        End Try

        SEQ_NUM = (Val(SEQ_NUM) + 1).ToString

        dr.Item("YEAR_ID") = _RYP
        dr.Item("VEND_CODE") = VEND_CODE
        dr.Item("CHECK_NUM") = ManualEntry
        dr.Item("SEQ_NUM") = SEQ_NUM
        dr.Item("INV_NUM") = ManualEntry
        dr.Item("INV_DATE") = DateTime.Now.ToShortDateString
        dr.Item("VOUCHER_NO") = ManualEntry
        dr.Item("AMT_PAID") = dd
        dr.Item("INV_AMT") = dd
        dr.Item("INV_1099_AMT") = dd
        dst.Tables("APT10992").Rows.Add(dr)

        Me.UpdateReported1099()

    End Sub

    Private Sub SetupPrintRecords()

        dst.Tables("APT10991rpt").Rows.Clear()
        Dim rowAPT10991rpt As DataRow = Nothing

        Dim sql As String = String.Empty

        Dim VEND_1099_BOX As Integer = 0

        For Each dr As DataRow In dst.Tables("APT10991").Select("", "VEND_CODE")
            If (dr.Item("VEND_TAX_ID") & String.Empty).ToString.Trim.Length <> 9 Then Continue For
            If Not "ES".Contains((dr.Item("VEND_TAX_ID_TYPE") & String.Empty).ToString.Trim) Then Continue For

            VEND_1099_BOX = Val(dr.Item("VEND_1099_BOX") & String.Empty)
            If (VEND_1099_BOX < 1 Or VEND_1099_BOX > 14 Or VEND_1099_BOX = 11 Or VEND_1099_BOX = 12) Then Continue For

            If Val(dr.Item("REPORTED_1099") & String.Empty) < _AMT_1099 Then Continue For

            rowAPT10991rpt = dst.Tables("APT10991rpt").NewRow
            For Each dc As DataColumn In dst.Tables("APT10991").Columns
                rowAPT10991rpt.Item(dc.ColumnName) = dr.Item(dc.ColumnName)
            Next
            dst.Tables("APT10991rpt").Rows.Add(rowAPT10991rpt)
        Next


    End Sub

#End Region

#Region "grdAPT10991"

    Private Sub grdAPT10991_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPT10991.AfterRowActivate
        Dim dvw As DataView = dst.Tables("APT10992").DefaultView
        dvw.RowFilter = "VEND_CODE = '" & grdAPT10991.ActiveRow.Cells("VEND_CODE").Value & "'"

        If Me.grdAPT10992.Rows.Count > 0 Then
            Me.grdAPT10992.ActiveRow = Me.grdAPT10992.Rows(0)
        End If

        Dim caption As String = "Payments made to Vendor: (" & grdAPT10991.ActiveRow.Cells("VEND_CODE").Value & ") " & grdAPT10991.ActiveRow.Cells("VEND_NAME").Value
        grdAPT10992.Text = caption

    End Sub

    Private Sub grdAPT10991_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPT10991.InitializeRow

        Dim VEND_1099_BOX As Integer = Val((e.Row.Cells("VEND_1099_BOX").Value & String.Empty).ToString.Trim)
        If (VEND_1099_BOX < 1 Or VEND_1099_BOX > 14 Or VEND_1099_BOX = 11 Or VEND_1099_BOX = 12) Then
            e.Row.Cells("VEND_1099_BOX").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("VEND_1099_BOX").Appearance.ForeColor = Drawing.Color.White
        Else
            e.Row.Cells("VEND_1099_BOX").Appearance.Reset()
        End If

        Dim VEND_TAX_ID As String = e.Row.Cells("VEND_TAX_ID").Value & String.Empty
        VEND_TAX_ID = VEND_TAX_ID.Replace(" ", "")

        If VEND_TAX_ID.Length <> 9 Then
            e.Row.Cells("VEND_TAX_ID").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("VEND_TAX_ID").Appearance.ForeColor = Drawing.Color.White
        Else
            e.Row.Cells("VEND_TAX_ID").Appearance.Reset()
        End If

        Dim VEND_TAX_ID_TYPE As String = e.Row.Cells("VEND_TAX_ID_TYPE").Value & String.Empty
        VEND_TAX_ID_TYPE = VEND_TAX_ID_TYPE.Replace(" ", "").Trim

        If Not (VEND_TAX_ID_TYPE = "E" Or VEND_TAX_ID_TYPE = "S") Then
            e.Row.Cells("VEND_TAX_ID_TYPE").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("VEND_TAX_ID_TYPE").Appearance.ForeColor = Drawing.Color.White
        Else
            e.Row.Cells("VEND_TAX_ID_TYPE").Appearance.Reset()
        End If

    End Sub

#End Region

#Region "grdAPT10992"

    Private Sub grdAPT10992_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPT10992.AfterRowsDeleted
        Me.UpdateReported1099()
    End Sub

    Private Sub grdAPT10992_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPT10992.AfterRowUpdate

        If Not dst.Tables.Contains("APTINVH1") Then Exit Sub

        Me.UpdateReported1099()

        Try
            Dim VoucherNo As String = e.Row.Cells("VOUCHER_NO").Value
            If VoucherNo = ManualEntry Then Exit Sub

            Dim dr As DataRow = dst.Tables("APTINVH1").Select("VOUCHER_NO = '" & VoucherNo & "'")(0)
            If Not dr Is Nothing Then
                dr.Item("INV_1099_AMT") = Val(e.Row.Cells("INV_1099_AMT").Value & String.Empty)
                If dr.Item("INV_1099_AMT") > 0 Then
                    dr.Item("INV_1099_IND") = "1"
                Else
                    dr.Item("INV_1099_IND") = "0"
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdAPT10992_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPT10992.BeforeRowUpdate

        If (e.Row.Cells("CHECK_NUM").Value & String.Empty).ToString.Trim = ManualEntry Then Exit Sub

        If Val(e.Row.Cells("INV_1099_AMT").Value & String.Empty) = 0 And Val(e.Row.Cells("AMT_PAID").Value & String.Empty) < 0 Then
            'Stop
        ElseIf Val(e.Row.Cells("INV_1099_AMT").Value & String.Empty) > Val(e.Row.Cells("AMT_PAID").Value & String.Empty) Then
            MsgBox("1099 Amount may not be more than the amount paid.", MsgBoxStyle.OkOnly, "Update Error")
            e.Cancel = True
        ElseIf Val(e.Row.Cells("INV_1099_AMT").Value & String.Empty) < 0 Then
            'MsgBox("1099 Amount may not be less than zero.", MsgBoxStyle.OkOnly, "Update Error")
            'e.Cancel = True
        End If
    End Sub

    Private Sub grdAPT10992_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdAPT10992.DoubleClickCell
        Try
            If Not e.Cell.Column.Key = "INV_1099_AMT" Then Exit Sub
            If Val(e.Cell.Value & String.Empty) = 0 Then
                e.Cell.Value = e.Cell.Row.Cells("INV_AMT").Value
            Else
                e.Cell.Value = 0
            End If
            grdAPT10992.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdAPT10992_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdAPT10992.BeforeRowsDeleted
        For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            If dr.Cells("CHECK_NUM").Value <> ManualEntry Then
                MsgBox("You can only delete manual entries. Set the 1099 payment to zero for all non Manual entries.", MsgBoxStyle.OkOnly, "Delete Error")
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub UpdateReported1099()

        Dim YEAR_ID As String = grdAPT10991.ActiveRow.Cells("YEAR_ID").Value
        Dim VEND_CODE As String = grdAPT10991.ActiveRow.Cells("VEND_CODE").Value

        Dim dd As Double = dst.Tables("APT10992").Compute("SUM(INV_1099_AMT)", "YEAR_ID = '" & YEAR_ID & "' AND VEND_CODE = '" & VEND_CODE & "'")
        grdAPT10991.ActiveRow.Cells("REPORTED_1099").Value = dd
    End Sub

#End Region
End Class