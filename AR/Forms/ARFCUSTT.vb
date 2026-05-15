Imports ABSolution

Public Class ARFCUSTT

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            MyBase.Create_TDA(.Tables.Add, "ARTCUSTT", "*")

        End With

        grdARTCUSTT.DataSource = dst.Tables("ARTCUSTT")
        Create_Summary(grdARTCUSTT, "CUST_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"

            Case "Update"

            Case "Update Customers"
                If MessageBox.Show("Do you want to convert Customer / Store Numbers?", "Update Customers", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ' Verify all from customer/stores exist
                Dim sql As String = String.Empty
                sql = "SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUSTT" _
                    & " MINUS " _
                    & " SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUST2"
                Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
                If tbl.Rows.Count > 0 Then
                    sql = String.Empty
                    For Each row As DataRow In tbl.Select("", "CUST_CODE, CUST_STORE_NO")
                        sql = row.Item("CUST_CODE") & " - " & row.Item("CUST_STORE_NO") & Environment.NewLine
                    Next

                    EMsg &= vbCr & "The following 'From' Customer / Stores do not Exist" & Environment.NewLine
                    EMsg &= sql
                End If

                ' Verify To Customers do not exist
                sql = "SELECT NEW_CUST_CODE, NEW_CUST_STORE_NO FROM ARTCUSTT" _
                    & " INTERSECT  " _
                    & " SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUST2"
                tbl = ASCDATA1.GetDataTable(sql)
                If tbl.Rows.Count > 0 Then
                    sql = String.Empty
                    For Each row As DataRow In tbl.Select("", "NEW_CUST_CODE, NEW_CUST_STORE_NO")
                        sql = row.Item("NEW_CUST_CODE") & " - " & row.Item("NEW_CUST_STORE_NO") & Environment.NewLine
                    Next

                    EMsg &= vbCr & "The following 'To' Customer / Stores already Exist" & Environment.NewLine
                    EMsg &= sql
                End If

                'Verify New Customer Code exists in customer master
                sql = "SELECT NEW_CUST_CODE FROM ARTCUSTT" _
                    & " MINUS " _
                    & " SELECT CUST_CODE FROM ARTCUST1"
                tbl = ASCDATA1.GetDataTable(sql)
                If tbl.Rows.Count > 0 Then
                    sql = String.Empty
                    For Each row As DataRow In tbl.Select("", "CUST_CODE")
                        sql = row.Item("CUST_CODE") & " - " & row.Item("CUST_STORE_NO") & Environment.NewLine
                    Next

                    EMsg &= vbCr & "The following 'To' Customer Code(s) do not exist in the Customer Master." & Environment.NewLine
                    EMsg &= sql
                End If

                ' Verify New customers have store numbers
                sql = "SELECT * FROM ARTCUSTT WHERE NEW_CUST_STORE_NO IS NULL"
                tbl = ASCDATA1.GetDataTable(sql)
                If tbl.Rows.Count > 0 Then
                    sql = String.Empty
                    For Each row As DataRow In tbl.Select("", "NEW_CUST_CODE, NEW_CUST_STORE_NO")
                        sql = row.Item("NEW_CUST_CODE") & " - " & row.Item("NEW_CUST_STORE_NO") & Environment.NewLine
                    Next

                    EMsg &= vbCr & "The following 'To' Customer / Stores already Exist" & Environment.NewLine
                    EMsg &= sql
                End If

                If EMsg.Length = 0 Then
                    If InputBox("Please provide the password.", "Update Customers") <> "Clifford" Then
                        EMsg &= vbCr & "Invalid password"
                    End If
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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode

                .Groups("Screen Control").Items("Update Customers").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update Customers").Visible = Not (ASCMAIN1.DBS_SERVER = ASCMAIN1.DBS_COMPANY)

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
            Load_Record()
        End If

        If ScreenMode Then
            grdARTCUSTT.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdARTCUSTT.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdARTCUSTT.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdARTCUSTT.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdARTCUSTT.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdARTCUSTT.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("ARTCUSTT").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)
        Fill_Records("ARTCUSTT", String.Empty, True, "select * From ARTCUSTT")
        Sort_grdColumns(grdARTCUSTT, "CUST_CODE,CUST_STORE_NO")
        grdARTCUSTT.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("ARTCUSTT")
            CommitTrans("Update Complete")
        Catch ex As Exception
            CommitTrans(ex.Message)
        End Try
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdARTCUSTT_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTT.BeforeRowUpdate

        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value & String.Empty

        e.Row.Cells("NEW_CUST_CODE").Value = e.Row.Cells("CUST_CODE").Value
        Dim NEW_CUST_CODE As String = e.Row.Cells("NEW_CUST_CODE").Value & String.Empty
        Dim NEW_CUST_STORE_NO As String = e.Row.Cells("NEW_CUST_STORE_NO").Value & String.Empty
        NEW_CUST_STORE_NO = NEW_CUST_STORE_NO.Trim

        If NEW_CUST_STORE_NO.Length = 0 Then
            MessageBox.Show("Missing 'To' Store number.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        NEW_CUST_STORE_NO = ASCMAIN1.Format_Field(NEW_CUST_STORE_NO, "CUST_STORE_NO")
        e.Row.Cells("NEW_CUST_STORE_NO").Value = NEW_CUST_STORE_NO

        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
        If rowARTCUST2 Is Nothing Then
            MessageBox.Show("Invalid 'From' Customer/Store, the combination must exist.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        e.Row.Cells("CUST_STORE_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty

        rowARTCUST2 = LookUp("ARTCUST2", New String() {NEW_CUST_CODE, NEW_CUST_STORE_NO})
        If rowARTCUST2 IsNot Nothing Then
            MessageBox.Show("Invalid 'To' Customer/Store, the combination cannot exist.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If


    End Sub

    Private Sub grdARTCUSTT_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTCUSTT.ClickCellButton

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim sql_where As String = String.Empty

        Select Case COLUMN_NAME
            Case "CUST_CODE"
                grdClickCellButton(grdARTCUSTT, sql_where)

            Case "CUST_STORE_NO"
                sql_where = "CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Value & "'"
                grdClickCellButton(grdARTCUSTT, sql_where)

            Case Else
                Exit Sub
        End Select

    End Sub

#End Region

#Region "Form Procedures"

    Private Sub UpdateCustomers()

        Try
            If 1 = 1 Then
                MessageBox.Show("Feature Not Turned on!", "Update Customers", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            BeginTrans()

            ASCMAIN1.Progress("")
            Dim sql As String = String.Empty

            ASCMAIN1.Progress("ARTCUST7")
            ASCDATA1.ExecuteSQL("DELETE FROM ARTCUST7 WHERE (CUST_CODE, CUST_STORE_NO) IN (SELECT CUST_CODE, NEW_CUST_STORE_NO FROM ARTCUSTT)")

            ASCMAIN1.Progress("SPTCWRXY")
            ASCDATA1.ExecuteSQL("DELETE FROM SPTCWRXY WHERE (CUST_CODE, CUST_STORE_NO) IN (SELECT CUST_CODE, NEW_CUST_STORE_NO FROM ARTCUSTT) " _
                        & " AND (CUST_CODE, CUST_STORE_NO) NOT IN (SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUSTT)")

            ASCMAIN1.Progress("SPTCWRXZ")
            ASCDATA1.ExecuteSQL("DELETE FROM SPTCWRXZ WHERE (CUST_CODE, CUST_STORE_NO) IN (SELECT CUST_CODE, NEW_CUST_STORE_NO FROM ARTCUSTT)  " _
                                & " AND (CUST_CODE, CUST_STORE_NO) NOT IN (SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUSTT)")

            ASCMAIN1.Progress("ARTCUSTT PL Sql")
            sql = " BEGIN DECLARE CURSOR C1 IS SELECT * FROM ARTCUSTT;"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= "                 UPDATE ARTCUST2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE ARTCUST7 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE ARTOPEN1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE ARTOPENX SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE GLTPROFS SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTBUDR1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTCUSTS SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTRETL1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTRETL2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTRETL4 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTRETL5 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE RSTRETLC SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SATAUTH1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SATAUTH2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SATSSUMS SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTDRBR0 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTINVH1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTINVH2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTORDR1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTORDR2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTRMAF1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SOTRTRN1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTCWRX2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTCWRXY SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTCWRXZ SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTCWXBD SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTCWXBI SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTDCOMB SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTMBUD1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTMXWS2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTMXWS3 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTMXWS4 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE SPTPYXI1 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE TATXREF4 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE TATXREF9 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= "                 UPDATE TATXREFC SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO WHERE CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            ASCMAIN1.Progress("ARTPYMT3")
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT ARTPYMT3.*, ARTPYMT2.CUST_CODE, ARTCUSTT.NEW_CUST_STORE_NO"
            sql &= " FROM ARTPYMT3,ARTPYMT2, ARTCUSTT"
            sql &= " WHERE ARTPYMT2.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO "
            sql &= " AND ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT3.PYMT_BATCH_LNO"
            sql &= " AND ARTPYMT2.CUST_CODE = ARTCUSTT.CUST_CODE "
            sql &= " AND ARTPYMT3.CUST_STORE_NO = ARTCUSTT.CUST_STORE_NO;"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " UPDATE ARTPYMT3 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO"
            sql &= "                                WHERE PYMT_BATCH_NO = R1.PYMT_BATCH_NO"
            sql &= "                                 AND PYMT_BATCH_LNO = R1.PYMT_BATCH_LNO"
            sql &= "                                 AND PYMT_BATCH_ILNO = R1.PYMT_BATCH_ILNO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            ASCMAIN1.Progress("RSTIMPR1")
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT RSTIMPR2.*, RSTIMPR1.CUST_CODE, ARTCUSTT.NEW_CUST_STORE_NO"
            sql &= " FROM RSTIMPR1,RSTIMPR2, ARTCUSTT"
            sql &= " WHERE RSTIMPR1.IMPORT_NO = RSTIMPR2.IMPORT_NO"
            sql &= " AND RSTIMPR1.CUST_CODE = ARTCUSTT.CUST_CODE "
            sql &= " AND RSTIMPR2.CUST_STORE_NO = ARTCUSTT.CUST_STORE_NO; "
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " UPDATE RSTIMPR2 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO"
            sql &= "                                WHERE IMPORT_NO = R1.IMPORT_NO"
            sql &= "                                 AND CUST_STORE_NO = R1.CUST_STORE_NO"
            sql &= "                                 AND CUST_DEPT_CODE = R1.CUST_DEPT_CODE;"
            sql &= " END LOOP; END; END"
            ASCDATA1.ExecuteSQL(sql)

            ASCMAIN1.Progress("SPTCOOP1")
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT SPTCOOPB.*, SPTCOOP1.CUST_CODE, ARTCUSTT.NEW_CUST_STORE_NO"
            sql &= " FROM SPTCOOPB, SPTCOOP1, ARTCUSTT"
            sql &= " WHERE SPTCOOP1.AUTH_NO = SPTCOOPB.AUTH_NO "
            sql &= " AND SPTCOOP1.CUST_CODE = ARTCUSTT.CUST_CODE "
            sql &= " AND SPTCOOPB.CUST_STORE_NO = ARTCUSTT.CUST_STORE_NO;"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " UPDATE SPTCOOPB SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO"
            sql &= "                                WHERE AUTH_NO = R1.AUTH_NO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            ASCMAIN1.Progress("SPTCOOP1")
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT SPTCOOP9.*, SPTCOOP1.CUST_CODE, ARTCUSTT.NEW_CUST_STORE_NO"
            sql &= " FROM SPTCOOP9, SPTCOOP1, ARTCUSTT"
            sql &= " WHERE SPTCOOP1.AUTH_NO = SPTCOOP9.AUTH_NO "
            sql &= " AND SPTCOOP1.CUST_CODE = ARTCUSTT.CUST_CODE "
            sql &= " AND SPTCOOP9.CUST_STORE_NO = ARTCUSTT.CUST_STORE_NO;"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " UPDATE SPTCOOP9 SET CUST_STORE_NO = R1.NEW_CUST_STORE_NO"
            sql &= "                                WHERE AUTH_NO = R1.AUTH_NO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            CommitTrans("Successful COnversion")

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            ASCMAIN1.Progress("")
        End Try
    End Sub


#End Region

End Class