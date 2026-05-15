Public Class APFVOID1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim tables As DataTableCollection = dst.Tables

        Create_TDA(tables.Add, "APTCHCK1", "*")
        Create_TDA(tables.Add, "APTCHCK2", "*", 2, False)
        Create_TDA(tables.Add, "APTVEND5", "*")
        Create_TDA(tables.Add, "APTINVH1", "*")

        Create_Lookup("GLTBANK1")
        Create_Lookup("APTVEND1")
        Create_Lookup("APTCHCK1")

        grdAPTCHCK1.DataSource = tables("APTCHCK1")

        Call Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count")
        Call Create_Summary(grdAPTCHCK1, "CHECK_AMT")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Proceed"
                Validate_Code("BANK_CODE")
                If Absx1.txtFor("VEND_CODE").Text <> "" Then

                    Validate_Code("VEND_CODE")
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("BANK_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Void Checks"
                If grdAPTCHCK1.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Checks Selected to be Voided"
                End If

            Case "Reverse Checks"
                If grdAPTCHCK1.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Checks Selected to be Reversed"
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

            Case "Proceed"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
                Absx1.CtlFor("CHECK_NUM").Focus()

            Case "Void Checks"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Reverse Checks"
                Stop ' not supported at this time
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        Dim items As UltraWinExplorerBar.UltraExplorerBarItemsCollection = _
         UltraExplorerBar1.Groups("Screen Control").Items

        items("Proceed").Settings.Enabled = not_iScreenMode
        items("Void Checks").Settings.Enabled = iScreenMode
        items("Reverse Checks").Settings.Enabled = iScreenMode
        items("Cancel").Settings.Enabled = iScreenMode

        items("Reverse Checks").Visible = False ' SOX issue

        'With UltraExplorerBar1
        '    .Groups("Screen Control").Items("Proceed").Settings.Enabled = not_iScreenMode
        '    .Groups("Screen Control").Items("Void Checks").Settings.Enabled = iScreenMode
        '    .Groups("Screen Control").Items("Reverse Checks").Settings.Enabled = iScreenMode
        '    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

        '    .Groups("Screen Control").Items("Reverse Checks").Visible = False ' SOX issue

        'End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        Absx1.txtFor("CHECK_NUM").ReadOnly = False

        chkStatus.Enabled = True

        grpCHECK_NUM.Visible = tf
        lblVendorOptional.Visible = Not tf
        chkStatus.Visible = tf
        chkCardView.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("BANK_CODE").Text = ""
        Absx1.txtFor("VEND_CODE").Text = ""
        Absx1.txtFor("CHECK_NUM").Text = ""

        dst.Tables("APTCHCK1").Rows.Clear()
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        dst.Tables("APTCHCK1").Rows.Clear()

        Absx1.txtFor("BANK_CODE").Text = HFs("BANK_CODE")
        Absx1.txtFor("VEND_CODE").Text = HFs("VEND_CODE")

        grdAPTCHCK1.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTCHCK1.DisplayLayout.Bands(0).SortedColumns.Add("CHECK_NUM", False)

    End Sub

    Sub Update_Record()
        BeginTrans()

        For Each rowAPTCHCK1 As DataRow In dst.Tables("APTCHCK1").Rows
            HFs("CHECK_NUM") = rowAPTCHCK1.Item("CHECK_NUM")
            HFs("VEND_CODE") = rowAPTCHCK1.Item("VEND_CODE")
            Dim CHECK_AMT As Double = rowAPTCHCK1.Item("CHECK_AMT")

            ' Update APTVEND5

            Dim rowAPTVEND5 As DataRow = Fill_Record("APTVEND5", New String() {HFs("VEND_CODE")}, True, False)
            rowAPTVEND5.Item("VEND_PAYMENTS_MTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_MTD") & "") - CHECK_AMT
            rowAPTVEND5.Item("VEND_PAYMENTS_YTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_YTD") & "") - CHECK_AMT
            rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") & "") - 1
            rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") & "") - 1
            Update_Record_TDA("APTVEND5")

            ' Update APTCHCK1
            rowAPTCHCK1.Item("REGISTER_IND_F") = "0"
            rowAPTCHCK1.Item("CHECK_STATUS") = "V"
            rowAPTCHCK1.Item("OPS_YYYYPP_F") = ASCMAIN1.CYP
            rowAPTCHCK1.Item("LAST_DATE") = DATETIME_STAMP
            rowAPTCHCK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            'If vr = "R" Then
            '    rowAPTCHCK1.Delete()
            'End If
            Update_Record_TDA("APTCHCK1")

            ' Update APTINVH1 via APTCHCK2
            Fill_Records("APTCHCK2", New String() {HFs("BANK_CODE"), HFs("CHECK_NUM")})
            For Each rowAPTCHCK2 As DataRow In dst.Tables("APTCHCK2").Rows
                Dim VOUCHER_NO As String = rowAPTCHCK2.Item("VOUCHER_NO")
                Dim INV_AMT_APPLIED As Double = Val(rowAPTCHCK2.Item("INV_AMT_APPLIED") & "")
                Dim INV_DISC_TAKEN As Double = Val(rowAPTCHCK2.Item("INV_DISC_TAKEN") & "")

                Dim rowAPTINVH1 As DataRow = Fill_Record("APTINVH1", New String() {VOUCHER_NO})
                rowAPTINVH1.Item("CHECK_NUM") = System.DBNull.Value
                rowAPTINVH1.Item("CHECK_DATE") = System.DBNull.Value
                rowAPTINVH1.Item("INV_LAST_PMT_DATE") = System.DBNull.Value ' SHOULD PROBABLY BE RECALCULATED

                If chkStatus.Checked Then
                    rowAPTINVH1.Item("INV_STATUS") = "H"
                Else
                    rowAPTINVH1.Item("INV_STATUS") = "O"
                End If

                rowAPTINVH1.Item("INV_PAYMENTS") = Val(rowAPTINVH1.Item("INV_PAYMENTS") & "") - INV_AMT_APPLIED
                rowAPTINVH1.Item("INV_DISC_TAKEN") = Val(rowAPTINVH1.Item("INV_DISC_TAKEN") & "") - INV_DISC_TAKEN
                rowAPTINVH1.Item("INV_BALANCE") = Val(rowAPTINVH1.Item("INV_BALANCE") & "") + INV_AMT_APPLIED

                rowAPTINVH1.Item("BATCH_NO_PYMT") = ""
                rowAPTINVH1.Item("BATCH_PYMT") = 0
                rowAPTINVH1.Item("BATCH_DISC") = 0

                Call Update_Record_TDA("APTINVH1")

                Call Write_Event_Log("APTINVH1", VOUCHER_NO, "Check " & HFs("CHECK_NUM") & " Voided")

                'If vr = "R" Then
                '    rowAPTCHCK2.Delete()
                'End If
            Next
        Next

        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Proceed", e)
                End If
            Case "CHECK_NUM"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Absx1.txtFor("CHECK_NUM").Text = ASCMAIN1.Format_Field(Absx1.txtFor("CHECK_NUM").Text, "CHECK_NUM")
                    HFs("CHECK_NUM") = Absx1.txtFor("CHECK_NUM").Text
                    Dim row As DataRow = LookUp("APTCHCK1", New String() {HFs("BANK_CODE"), HFs("CHECK_NUM")})

                    row = ASCDATA1.GetDataRow("SELECT * FROM APTCHCK1 WHERE BANK_CODE = :PARM1 AND CHECK_NUM = :PARM2", "VV", New String() {HFs("BANK_CODE"), HFs("CHECK_NUM")})

                    If row IsNot Nothing Then
                        If HFs("VEND_CODE") <> "" And HFs("VEND_CODE") <> HFs("VEND_CODE") Then
                            MsgBox("Check " & HFs("CHECK_NUM") & " was not issued to Vendor " & HFs("VEND_CODE"), MsgBoxStyle.OkOnly, "Cannot Perform Action")
                        Else
                            If row.Item("CHECK_STATUS") = "V" Then
                                MsgBox("Check " & HFs("CHECK_NUM") & " has already been Voided", MsgBoxStyle.OkOnly, "Cannot Perform Action")
                            ElseIf row.Item("BATCH_NO_CLEARED") & "" <> "" Then
                                MsgBox("Check " & HFs("CHECK_NUM") & " has already been Cleared", MsgBoxStyle.OkOnly, "Cannot Perform Action")

                            Else

                                Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").Rows.Find(New String() {HFs("BANK_CODE"), HFs("CHECK_NUM")})
                                If rowAPTCHCK1 Is Nothing Then
                                    dst.Tables("APTCHCK1").LoadDataRow(row.ItemArray, True)
                                End If
                            End If
                        End If
                        Absx1.txtFor("CHECK_NUM").Text = ""
                    Else
                        Absx1.txtFor("CHECK_NUM").Text = ""
                    End If
                End If
        End Select
    End Sub
#End Region

    Private Sub chkCardView_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCardView.CheckedChanged
        grdAPTCHCK1.DisplayLayout.Bands(0).CardView = chkCardView.Checked
    End Sub

    Private Sub UltraGroupBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox1.Click

    End Sub
End Class