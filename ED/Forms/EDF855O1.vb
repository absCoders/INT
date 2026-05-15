Imports ABSolution

Public Class EDF855O1

    Const Accept As String = "AK"
    Const Reject As String = "RJ"
    Const AcceptwithChange As String = "AH"

    Private Const EDI_PROCESS_IND As String = "1"
    Private Const EDI_APPLICATION_ID As String = "PR"

    Private Enum DocTypes
        EDI850
        EDI860
    End Enum

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = " Select EDI_DOC_SEQ_NO, '850' EDI_TYPE, CUST_CODE, EDI_DEPARTMENT, EDI_PO_DATE, EDI_PO_NO, EDI_STORE," _
                & " EDI_START_DATE, EDI_END_DATE, EDI_PO_PURP, EDI_PO_TYPE" _
                & " FROM EDT850T1"
            Create_TDA(.Tables.Add, "EDT850TX", ASCMAIN1.sql, 0, False, String.Empty, 2)

            Create_TDA(.Tables.Add, "EDT850T1", "*")
            Create_TDA(.Tables.Add, "EDT850T2", "*")
            Create_TDA(.Tables.Add, "EDT850T3", "*")
            Create_TDA(.Tables.Add, "EDT850T4", "*")

            Create_TDA(.Tables.Add, "EDT860T1", "*")
            Create_TDA(.Tables.Add, "EDT860T2", "*")
            Create_TDA(.Tables.Add, "EDT860T3", "*")
            Create_TDA(.Tables.Add, "EDT860T4", "*")

            dst.Tables("EDT850TX").Columns.Add("EDI_ACK_TYPE", GetType(System.String))
            dst.Tables("EDT850TX").Columns("EDI_ACK_TYPE").DefaultValue = "00"
            dst.Tables("EDT850TX").Columns.Add("CUST_NAME", GetType(System.String), "CUST_CODE")

            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT855O1", "*")
            Create_TDA(.Tables.Add, "EDTTRPM1", "*")
            Create_TDA(.Tables.Add, "EDTTRPM2", "*")
            Create_TDA(.Tables.Add, "EDTTRPM3", "*")

            ASCMAIN1.sql = "Select * from EDTTRPM1 where EDI_DOC_NO = '855'"
            Fill_Records("EDTTRPM1", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from EDTTRPM2 where EDI_DOC_NO = '855'"
            Fill_Records("EDTTRPM2", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from EDTTRPM3 where EDI_DOC_NO = '855'"
            Fill_Records("EDTTRPM3", String.Empty, True, ASCMAIN1.sql)

        End With

        grdEDT850TX.DataSource = dst.Tables("EDT850TX")

        grdEDT850T2.DataSource = dst.Tables("EDT850T2")
        grdEDT850T3.DataSource = dst.Tables("EDT850T3")
        grdEDT850T4.DataSource = dst.Tables("EDT850T4")

        grdEDT860T2.DataSource = dst.Tables("EDT860T2")
        grdEDT860T3.DataSource = dst.Tables("EDT860T3")
        grdEDT860T4.DataSource = dst.Tables("EDT860T4")

        For Each gridColumn As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdEDT850TX.DisplayLayout.Bands(0).Columns
            If gridColumn.Key = "EDI_ACK_TYPE" Then
                gridColumn.CellActivation = UltraWinGrid.Activation.AllowEdit
                gridColumn.CellAppearance.BackColor = Drawing.Color.LightBlue
            Else
                gridColumn.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        grdEDT850TX.DisplayLayout.UseFixedHeaders = True
        grdEDT850TX.DisplayLayout.Bands(0).Columns("EDI_ACK_TYPE").Header.Fixed = True
        grdEDT850TX.DisplayLayout.Bands(0).Columns("EDI_DOC_SEQ_NO").Header.Fixed = True

        Dim vlProcessInd As String() = {":", "00:Not Processed", Accept & ":Accept", Reject & ":Reject", AcceptwithChange & ":Accept With Change"}

        ASCMAIN1.Add_Value_List(grdEDT850TX, "EDI_ACK_TYPE", , vlProcessInd)
        ASCMAIN1.Add_Value_List(grdEDT850TX, "CUST_NAME", "SELECT CUST_CODE, CUST_NAME FROM ARTCUST1")
        ASCMAIN1.Add_Value_List(grdEDT850TX, "EDI_TYPE", , New String() {":", "850:P.O.", "860:P.O. Change"})

        tab850details.Dock = DockStyle.Fill
        tab860details.Dock = DockStyle.Fill

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel your changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Update"

                Dim edi850 As Int32 = dst.Tables("EDT850TX").Select("EDI_ACK_TYPE <> '00' AND EDI_TYPE = '850'").Length
                Dim edi860 As Int32 = dst.Tables("EDT850TX").Select("EDI_ACK_TYPE <> '00' AND EDI_TYPE = '860'").Length

                If edi850 = 0 AndAlso edi860 = 0 Then
                    EMsg &= vbCr & "All records set to 'Not Processed' - Nothing to Update."
                    Exit Select
                End If

                Dim zmsg As String = "There are " & edi850 & " P.O. records modified." & Environment.NewLine _
                                     & "There are " & edi860 & " P.O. Changes records modified." & Environment.NewLine _
                                     & Environment.NewLine _
                                     & "Do you want to continue?"

                If MessageBox.Show(zmsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        splEDF855O1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        For Each tableName As String In New String() {"EDT850T1", "EDT850T2", "EDT850T3", "EDT850T4", _
                                                       "EDT860T1", "EDT860T2", "EDT860T3", "EDT860T4", _
                                                       "EDT855O1", "EDTSYSIH", "EDT850TX"}
            dst.Tables(tableName).Rows.Clear()
        Next

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading data")

        EnforceConstraints(False)

        'AND EDI_PO_TYPE IN ('KC', 'KB', 'BK')" _
        ASCMAIN1.sql = $"SELECT * FROM EDT850T1 WHERE NVL(EDI_PROCESS_IND,'0') = '0' AND EDI_PO_TYPE IN ('KC', 'KB', 'BK')
                        AND CUST_CODE IN 
                        (
                            Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '855' 
                            Union 
                            Select CUST_CODE from EDTTRPM2 where EDI_DOC_NO = '855'
                            Union 
                            Select CUST_CODE from EDTTRPM3 where EDI_DOC_NO = '855'
                        ) AND COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'"
        Dim temptable850 As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & temptable850 & ")"
        Fill_Records("EDT850T1", String.Empty, True, ASCMAIN1.sql)
        Fill_Records("EDT850T2", String.Empty, True, ASCMAIN1.sql.Replace("EDT850T1", "EDT850T2"))
        Fill_Records("EDT850T3", String.Empty, True, ASCMAIN1.sql.Replace("EDT850T1", "EDT850T3"))
        Fill_Records("EDT850T4", String.Empty, True, ASCMAIN1.sql.Replace("EDT850T1", "EDT850T4"))

        ' AND EDI_PO_TYPE IN ('KC', 'KB', 'BK')" _
        ASCMAIN1.sql = $"SELECT * FROM EDT860T1 WHERE NVL(EDI_PROCESS_IND,'0') = '0' AND EDI_PO_TYPE IN ('KC', 'KB', 'BK')
                        AND CUST_CODE IN 
                        (
                            Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '855' 
                            Union 
                            Select CUST_CODE from EDTTRPM2 where EDI_DOC_NO = '855'
                            Union 
                            Select CUST_CODE from EDTTRPM3 where EDI_DOC_NO = '855'
                        ) 
                        AND COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'"
        Dim temptable860 As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & temptable860 & ")"
        Fill_Records("EDT860T1", String.Empty, True, ASCMAIN1.sql)
        Fill_Records("EDT860T2", String.Empty, True, ASCMAIN1.sql.Replace("EDT860T1", "EDT860T2"))
        Fill_Records("EDT860T3", String.Empty, True, ASCMAIN1.sql.Replace("EDT860T1", "EDT860T3"))
        Fill_Records("EDT860T4", String.Empty, True, ASCMAIN1.sql.Replace("EDT860T1", "EDT860T4"))

        ASCMAIN1.sql = "Select EDI_DOC_SEQ_NO, '850' EDI_TYPE, CUST_CODE, EDI_DEPARTMENT, EDI_PO_DATE, EDI_PO_NO, EDI_STORE," _
        & " EDI_START_DATE, EDI_END_DATE, EDI_PO_PURP, EDI_PO_TYPE" _
        & " FROM EDT850T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & temptable850 & ")" _
        & " UNION" _
        & " Select EDI_DOC_SEQ_NO, '860' EDI_TYPE, CUST_CODE, EDI_DEPARTMENT, EDI_ORIG_DATE, EDI_PO_NO, EDI_STORE_NO," _
        & " EDI_START_DATE, EDI_END_DATE, EDI_TRANS_PURP, EDI_PO_TYPE" _
        & " FROM EDT860T1 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & temptable860 & ")"
        Fill_Records("EDT850TX", String.Empty, True, ASCMAIN1.sql)

        EnforceConstraints(True)

        grdEDT850TX.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdEDT850T2.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdEDT850T3.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdEDT850T4.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()
        Try
            Dim EDI_OUTBOUND_DOC_NO As String = String.Empty
            Dim EDI_OUR_ID As String = String.Empty
            Dim EDI_TP_ID As String = String.Empty
            Dim ediApplicationID As String = String.Empty
            Dim EDI_STATUS As String = String.Empty
            Dim CUST_CODE As String = String.Empty
            Dim rowEDTTRPM1 As DataRow = Nothing
            Dim ACK_TYPE As String = String.Empty

            BeginTrans()

            For Each rowEDT850TX As DataRow In dst.Tables("EDT850TX").Select("EDI_ACK_TYPE <> '00'")
                Dim EDI_DOC_SEQ_NO As String = rowEDT850TX.Item("EDI_DOC_SEQ_NO") & String.Empty
                Dim EDI_TYPE As String = rowEDT850TX.Item("EDI_TYPE") & String.Empty

                Select Case EDI_TYPE

                    Case "850"
                        Dim rowEDT850T1 As DataRow = dst.Tables("EDT850T1").Rows.Find(New Object() {EDI_DOC_SEQ_NO})
                        CUST_CODE = rowEDT850T1.Item("CUST_CODE") & String.Empty
                        If dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        ElseIf dst.Tables("EDTTRPM2").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM2").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        ElseIf dst.Tables("EDTTRPM3").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM3").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        Else
                            Continue For
                        End If

                        EDI_TP_ID = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

                        If dst.Tables("EDTTRPM1").Select("EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("EDI_TP_ID = '" & EDI_TP_ID & "'")(0)
                        Else
                            MessageBox.Show($"Cannot determine EDI Status for Customer {CUST_CODE}, Trading Partner ID: {EDI_TP_ID}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Continue For
                        End If

                        EDI_OUR_ID = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
                        EDI_STATUS = rowEDTTRPM1.Item("EDI_STATUS") & String.Empty

                        ACK_TYPE = rowEDT850TX.Item("EDI_ACK_TYPE") & String.Empty

                        If Not (ACK_TYPE = Accept OrElse ACK_TYPE = Reject OrElse ACK_TYPE = AcceptwithChange) Then
                            Continue For
                        End If

                        EDI_OUTBOUND_DOC_NO = CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, EDI_STATUS)
                        Create855O1Record(EDI_OUTBOUND_DOC_NO, ACK_TYPE, rowEDT850T1, DocTypes.EDI850)

                        rowEDT850T1.Item("EDI_PROCESS_IND") = "1"

                    Case "860"

                        Dim rowEDT860T1 As DataRow = dst.Tables("EDT860T1").Rows.Find(New Object() {EDI_DOC_SEQ_NO})

                        CUST_CODE = rowEDT860T1.Item("CUST_CODE") & String.Empty
                        If dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        ElseIf dst.Tables("EDTTRPM2").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM2").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        ElseIf dst.Tables("EDTTRPM3").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM3").Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                        Else
                            Continue For
                        End If


                        EDI_TP_ID = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty
                        If dst.Tables("EDTTRPM1").Select("EDI_TP_ID = '" & EDI_TP_ID & "'").Length > 0 Then
                            rowEDTTRPM1 = dst.Tables("EDTTRPM1").Select("EDI_TP_ID = '" & EDI_TP_ID & "'")(0)
                        Else
                            MessageBox.Show($"Cannot determine EDI Status for Customer {CUST_CODE}, Trading Partner ID: {EDI_TP_ID}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Continue For
                        End If

                        EDI_OUR_ID = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
                        EDI_STATUS = rowEDTTRPM1.Item("EDI_STATUS") & String.Empty

                        ACK_TYPE = rowEDT850TX.Item("EDI_ACK_TYPE") & String.Empty

                        If Not (ACK_TYPE = Accept OrElse ACK_TYPE = Reject OrElse ACK_TYPE = AcceptwithChange) Then
                            Continue For
                        End If

                        EDI_OUTBOUND_DOC_NO = CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, EDI_STATUS)
                        Create855O1Record(EDI_OUTBOUND_DOC_NO, ACK_TYPE, rowEDT860T1, DocTypes.EDI860)

                        rowEDT860T1.Item("EDI_PROCESS_IND") = "1"

                End Select

            Next

            Update_Record_TDA("EDTSYSIH")
            Update_Record_TDA("EDT855O1")
            Update_Record_TDA("EDT850T1")
            Update_Record_TDA("EDT860T1")

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Private Function CreateEDTSYSIH(ByVal EDI_OUR_ID As String, _
                                    ByVal EDI_TP_ID As String, _
                                    ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
        rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
        rowEDTSYSIH.Item("EDI_APPLICATION_ID") = EDI_APPLICATION_ID
        If EDI_STATUS = "P" Then
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = EDI_PROCESS_IND
        Else
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "T"
        End If
        rowEDTSYSIH.Item("EDI_OUR_ID") = EDI_OUR_ID
        rowEDTSYSIH.Item("EDI_TP_ID") = EDI_TP_ID
        rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
        rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
        dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = EDI_OUTBOUND_DOC_NO

    End Function

    Private Sub Create855O1Record(ByVal EDI_OUTBOUND_DOC_NO As String, ByVal Ack_Type As String, rowData As DataRow, ByVal inDocType As DocTypes)
        Dim rowEDT855O1 As DataRow = dst.Tables("EDT855O1").NewRow
        With rowEDT855O1
            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("ORDR_CUST_PO") = rowData.Item("EDI_PO_NO")

            If inDocType = DocTypes.EDI860 Then
                .Item("ORDR_PO_DATE") = rowData.Item("EDI_CHANGE_DATE")
                .Item("EDI_PURPOSE_CODE") = rowData.Item("EDI_TRANS_PURP")
            Else
                .Item("ORDR_PO_DATE") = rowData.Item("EDI_PO_DATE")
                .Item("EDI_PURPOSE_CODE") = rowData.Item("EDI_PO_PURP")
            End If

            .Item("ORDR_SHIP_DATE") = rowData.Item("EDI_START_DATE")
            .Item("ORDR_CANCEL_DATE") = rowData.Item("EDI_END_DATE")
            .Item("EDI_SUPPLIER_NO") = rowData.Item("EDI_SUPPLIER_NO")
            .Item("EDI_DEPARTMENT") = rowData.Item("EDI_DEPARTMENT")
            .Item("REQUEST_DATE") = Format(Now, "dd-MMM-yy")
            .Item("EDI_ACK_TYPE") = Ack_Type
            .Item("AS_OF_DATE") = Format(Now, "dd-MMM-yy")
            .Item("INIT_DATE") = Format(Now, "dd-MMM-yy")
            .Item("INIT_OPER") = ASCMAIN1.USER_ID

            '.Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
            '.Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
            '.Item("EDI_ARRIVAL_DATE") = rowSOTORDR1.Item("EDI_ARRIVAL_DATE")
        End With
        dst.Tables("EDT855O1").Rows.Add(rowEDT855O1)
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdEDT850TX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT850TX.AfterRowActivate

        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim EDI_TYPE As String = String.Empty

        If grdEDT850TX.ActiveRow IsNot Nothing AndAlso Not grdEDT850TX.ActiveRow.IsAddRow Then
            EDI_DOC_SEQ_NO = grdEDT850TX.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
            EDI_TYPE = grdEDT850TX.ActiveRow.Cells("EDI_TYPE").Value & String.Empty
        End If

        Select Case EDI_TYPE

            Case "850"
                tab850details.Visible = True
                tab860details.Visible = False

                Dim dvw2 As DataView = DirectCast(grdEDT850T2.DataSource, DataTable).DefaultView
                dvw2.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw2.Sort = "EDI_DTL_SEQ"
                grdEDT850T2.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

                Dim dvw3 As DataView = DirectCast(grdEDT850T3.DataSource, DataTable).DefaultView
                dvw3.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw3.Sort = "EDI_DTL_SEQ, EDI_SDQ_SEQ"
                grdEDT850T3.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

                Dim dvw4 As DataView = DirectCast(grdEDT850T4.DataSource, DataTable).DefaultView
                dvw4.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw4.Sort = "EDI_CMT_SEQ"
                grdEDT850T4.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

            Case "860"
                tab850details.Visible = False
                tab860details.Visible = True

                Dim dvw2 As DataView = DirectCast(grdEDT860T2.DataSource, DataTable).DefaultView
                dvw2.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw2.Sort = "EDI_ORIG_DTL_SEQ"
                grdEDT860T2.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

                Dim dvw3 As DataView = DirectCast(grdEDT860T3.DataSource, DataTable).DefaultView
                dvw3.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw3.Sort = "EDI_ORIG_DTL_SEQ, EDI_SDQ_SEQ"
                grdEDT860T3.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

                Dim dvw4 As DataView = DirectCast(grdEDT860T4.DataSource, DataTable).DefaultView
                dvw4.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                dvw4.Sort = "EDI_CMT_SEQ"
                grdEDT860T4.Text = "EDI DOC SEQUENCE: " & EDI_DOC_SEQ_NO

            Case Else
                tab850details.Visible = False
                tab860details.Visible = False
        End Select
    End Sub

#End Region

End Class