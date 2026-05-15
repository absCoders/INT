Public Class ARFCCTR1

    Dim dt As DataTable = Nothing
    Dim FILENAME As String
    Dim rowARTCCTR1 As DataRow

    Dim NO_UPDATE As Boolean = False
    Private isProcessingPayPal As Boolean = False
    Private Const numCCDays As Integer = -4


    ' Note on AR Item Matching Process"
    ' When a Transation Batch is imported from spreadsheet, 
    ' SOTINVH1 records are matched based on ORDR_CUST_PO, and then on ORDR_NO_WEB (see Load Record)
    ' the match will look for Invoices if the TRANS_AMT is >=0, else it will look for Credits
    ' whatever record is found, the user may swtich to a different AR item in the Deposit process
    ' if no AR record is matched, and the transaction is selected in the deposit, a non-AR cash receipt record will be created 


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ARTCCTR1", "*")
            Create_TDA(.Tables.Add, "ARTCCTR2", "*", 1, , , , "CUST_CODE,INV_TYPE,INV_NUM,INV_TOTAL_AMOUNT,INV_BALANCE,INV_FREIGHT,PYMT_BATCH_NO,PYMT_BATCH_LNO")
            With .Tables("ARTCCTR2")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("INV_BALANCE_SEL", GetType(System.Decimal), "IIF(SEL='1',INV_BALANCE,0)")
                .Columns.Add("TRANS_AMT_SEL", GetType(System.Decimal), "IIF(SEL='1',TRANS_AMT,0)")
                '.Columns.Add("INV_NUM")
                '.Columns.Add("INV_TOTAL_AMOUNT", GetType(System.Decimal))
                '.Columns.Add("INV_BALANCE", GetType(System.Decimal))
                .Columns.Add("INV_BALANCE_NEW", GetType(System.Decimal), "INV_BALANCE - TRANS_AMT")
                '.Columns.Add("INVAILD_TRANS_AMT", GetType(System.String))
                .Columns.Add("MISSING_ORDER_NUMBER", GetType(System.String))
            End With

            ASCMAIN1.sql = "SELECT ARTCCTR1.* from ARTCCTR1 where ARTCCTR1.STATUS = '0'"
            Create_TDA(.Tables.Add, "ARTCCTRX", "**", 0, False, "VVV", 1)

            With .Tables.Add("ARTCCTRT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")
            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
            Create_TDA(.Tables.Add, "ARTPYMT4", "*")
            Create_TDA(.Tables.Add, "ARTPYMT5", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add("ARTOPENX"), "ARTOPEN1", "*")

        End With

        grdARTCCTRX.DataSource = dst.Tables("ARTCCTRX")
        grdARTCCTR2.DataSource = dst.Tables("ARTCCTR2")
        grdARTCCTRT.DataSource = dst.Tables("ARTCCTRT")
        grdARTOPENX.DataSource = dst.Tables("ARTOPENX")

        Create_Summary(grdARTCCTRX, "CCTB_NO", "Count")
        Create_Summary(grdARTCCTRX, New String() {"TOTAL_SLS", "TOTAL_RTN", "TOTAL_TRX"})

        Create_Summary(grdARTCCTR2, "CCTB_LNO", "Count")
        Create_Summary(grdARTCCTR2, New String() {"TRANS_AMT", "INV_TOTAL_AMOUNT", "INV_BALANCE", "INV_BALANCE_NEW", "SEL", "TRANS_AMT_SEL", "INV_BALANCE_SEL", "SERVICE_FEE", "TRANS_AMT_NET"})

        With grdARTCCTR2.DisplayLayout.Bands(0)
            .Columns("CCTB_NO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CCTB_NO").Header.Fixed = True
            .Columns("CCTB_LNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CCTB_LNO").Header.Fixed = True
            For Each COLUMN_NAME As String In New String() {"SEL", "INV_BALANCE_SEL", "TRANS_AMT_SEL"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightGoldenrodYellow
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "INV_TYPE", "INV_NUM", "INV_TOTAL_AMOUNT", "INV_FREIGHT", "INV_BALANCE", "INV_BALANCE_NEW"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightCyan
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"ORDER_NUMBER", "TRANS_AMT"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            '  .Columns("CCTB_LNO").Hidden = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"SEL"}.Contains(gcol.Key) Then ' "CUST_CODE", "INV_TYPE", "INV_NUM", "SEL"
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        grpDeposit.Left = grpCCTrans.Left
        grpDeposit.Top = grpCCTrans.Top
        grpDeposit.Dock = DockStyle.Right

        chkBestMatch.Left = grpCCTrans.Left
        chkBestMatch.Top = Absx1.txtFor("BANK_CODE").Top


        chkUnappliedOnly.Appearance.ForeColor = System.Drawing.Color.White
        chkUnappliedOnly.Appearance.BackColor = System.Drawing.Color.FromArgb(98, 160, 232)
        chkUnappliedOnly.Appearance.BackColor2 = System.Drawing.Color.FromArgb(83, 115, 191)
        chkUnappliedOnly.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Start Deposit"
                Validate_Code("BANK_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("F", "ARFPYMT2") Then Exit Sub
                End If

            Case "Import Cynergy", "Import PayPal"
                Validate_Code("BANK_CODE")
                isProcessingPayPal = False

                If EMsg = "" Then
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If InStr("AR", rowGLTBANK1.Item("BANK_USE") & "") = 0 Then
                        EMsg &= vbCr & "Bank " & Absx1.txtFor("BANK_CODE").Text & " is not defined for use in A/R"
                    End If
                End If

                FILENAME = String.Empty
                If EMsg = "" AndAlso eItemKey = "Import PayPal" Then

                    isProcessingPayPal = True
                    GetPayPalData()
                Else
                    dt = Gembox_Import_Sheet_to_DataTable(0, FILENAME)
                    If dt Is Nothing Then
                        EMsg &= vbCr & "You Must Select a Cynergy Transaction Export file to Create a Batch"
                    Else
                        If dt.Rows.Count = 0 Or dt.Columns.Count = 0 Then
                            EMsg &= vbCr & "Error Loading Export file"
                        End If
                    End If
                End If

            Case "Edit"
                Validate_Code("CCTB_NO")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCCTR1", Absx1.txtFor("CCTB_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                If grdARTCCTR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Transactions Entered into Batch"
                End If

                If dst.Tables("ARTCCTR2").Select("ISNULL(INV_BALANCE,0) <> ISNULL(TRANS_AMT,0)").Length <> 0 Then
                    If Not chkAllow.Checked Then
                        EMsg &= vbCr & "Some Transactions could not be reconciled to Open AR Items"
                        chkAllow.Visible = True
                    End If
                End If

            Case "Update Deposit"

                If Absx1.dteFor("PYMT_BATCH_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Deposit Date is Required"
                Else
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)

                    If Format(CDate(Absx1.dteFor("PYMT_BATCH_DATE").Value), "yyyyMMdd") > Format(CDate(rowGLTPARM2.Item("PRD_END_DATE")), "yyyyMMdd") Then
                        EMsg &= vbCr & "Deposit Date " & Absx1.txtFor("BANK_CODE").Text & " is later than Current Period End Date"
                    Else
                        ' 02/03/14 AMIRAH REQ THAT SHE BE ALLOWED TO POST JAN INTO FEB, BUT NOT FEB INTO JAN
                        'rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                        'If Format(CDATE(Absx1.dteFor("PYMT_BATCH_DATE").Value), "yyyyMMdd") <= Format(CDATE(rowGLTPARM2.Item("PRD_END_DATE")), "yyyyMMdd") Then
                        '    EMsg &= vbCr & "Deposit Date " & Absx1.txtFor("BANK_CODE").Text & " is earlier than Current Period Start Date " & Format(CDate(rowGLTPARM2.Item("PRD_END_DATE")).AddDays(1), "MM/dd/yyyy")
                        'End If
                    End If
                End If


                If EMsg = "" Then
                    Dim TOTAL_CC As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(TRANS_AMT)", "SEL = '1'") & "")
                    Dim TOTAL_AR As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(INV_BALANCE)", "SEL = '1'") & "")
                    Dim CC_COUNT As Int64 = Val(dst.Tables("ARTCCTR2").Compute("COUNT(CCTB_NO)", "SEL = '1'") & "")
                    Dim AR_CHANGED_COUNT As Int64 = Val(dst.Tables("ARTCCTR2").Compute("COUNT(CCTB_NO)", "SEL = '1' AND ISNULL(INV_TOTAL_AMOUNT,0)<>ISNULL(INV_BALANCE,0)") & "")
                    Dim CC_NO_AR As Int64 = Val(dst.Tables("ARTCCTR2").Compute("COUNT(CCTB_NO)", "SEL = '1' AND ISNULL(INV_NUM,'?') = '?'") & "")
                    Dim CC_FEE As Decimal = Val(Absx1.numFor("CC_FEE").Value & "")
                    Dim DEPOSIT_AMT As Decimal = Val(Absx1.numFor("DEPOSIT_AMT").Value & "")
                    Dim TOTAL_DEPOSIT As Decimal = TOTAL_CC - CC_FEE
                    If TOTAL_DEPOSIT <> DEPOSIT_AMT Then

                        If MsgBox("Credit Card Charges less Fee does not agree with Deposit Amount entered." _
                               & vbCrLf & vbCrLf & "If you proceed, some payments will be recorded as unapplied," _
                               & vbCrLf & " and you will have to apply them manually." _
                               & vbCrLf & vbCrLf & "Do you want to Proceed?", _
                               MsgBoxStyle.YesNo, "Verification to Proceed") = vbNo Then
                            Exit Sub
                        End If

                    End If

                    Dim msg As String = "" _
                                        & "You have created a Deposit for " & Format(TOTAL_DEPOSIT, "$#,##0.00") _
                                        & vbCrLf & " and have selected " & CStr(CC_COUNT) & " Credit Card transactions totaling " & Format(TOTAL_CC, "$#,##0.00") & "." _
                                        & vbCrLf & vbCrLf & "These CC transactions are associated with Open AR items totaling " & Format(TOTAL_AR, "$#,##0.00") & "." _
                                        & vbCrLf & "There are " & CStr(CC_NO_AR) & " CC transactions which are NOT associated with an Open AR Item." _
                                        & vbCrLf & "There are " & CStr(AR_CHANGED_COUNT) & " AR Items with an current balance <> the original invoice amount." _
                                        & vbCrLf & "There is a Bank Fee for " & Format(CC_FEE, "$#,##0.00") & "." _
                                        & vbCrLf & vbCrLf & "If you continue, the Open AR Items will be keyed off as Paid" _
                                        & vbCrLf & " and the Deposit will be recorded as a Cash Receipt." _
                                        & vbCrLf & vbCrLf & "Continue with Update?"
                    If MsgBox(msg, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

                End If

            Case "Delete"
                If MsgBox("Are you sure you want to Delete?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Start Deposit"
                EntryMode = "S"
                Mode_Settings(True)
                Refresh_Open_AR()
                Display_Totals()
                Absx1.dteFor("PYMT_BATCH_DATE").Value = DATETIME_STAMP.Date


            Case "Import Cynergy", "Import PayPal"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Update Deposit"
                Update_Deposit()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Edit List"
                Print_Report()

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '"Import Cynergy", "Import PayPal"

                    .Items("Import Cynergy").Settings.Enabled = not_iScreenMode
                    .Items("Import PayPal").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = Not This_Record_Inquiry_Only
                    .Items("Done").Visible = This_Record_Inquiry_Only

                    .Items("Update").Visible = ScreenMode And Not This_Record_Inquiry_Only

                    If NO_UPDATE Then
                        .Items("Update").Visible = False
                    End If

                    .Items("Cancel").Visible = ScreenMode And Not This_Record_Inquiry_Only
                    .Items("Delete").Visible = ScreenMode And Not This_Record_Inquiry_Only And (EntryMode = "E")

                End With
                With .Groups("Credit Card Deposit")
                    .Items("Start Deposit").Settings.Enabled = not_iScreenMode
                    .Items("Update Deposit").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Edit List").Settings.Enabled = iScreenMode

                End With

                .Groups("Totals").Visible = ScreenMode
            End With

            Setup_tabCC()

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpDeposit, False)

        If Not This_Record_Inquiry_Only Then
            Absx1.dteFor("CCTB_DATE").ReadOnly = False
        End If
        chkAllow.Visible = False
        chkAllow.Checked = False
        lblCCTB_MEMO.Visible = ScreenMode And EntryMode <> "S"
        Absx1.txtFor("CCTB_MEMO").Visible = ScreenMode And EntryMode <> "S"
        lblCCTB_DATE.Visible = ScreenMode And EntryMode <> "S"
        Absx1.dteFor("CCTB_DATE").Visible = ScreenMode And EntryMode <> "S"

        ' grdARTCCTR2.Visible = ScreenMode
        tabCC.Visible = Not ScreenMode Or EntryMode = "S"

        If ScreenMode Then
            If EntryMode = "S" Then
                Setup_grdARTCCTR2(False)
                Sort_grdColumns(grdARTCCTR2, "CCTB_NO,CCTB_LNO")
                grdARTCCTR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Else
                grdARTCCTR2.Parent = spl.Panel2
                grdARTCCTR2.DisplayLayout.GroupByBox.Hidden = True
                Show_Filter(grdARTCCTR2, False)
                Setup_grdARTCCTR2(True)
                With grdARTCCTR2.DisplayLayout.Override
                    If This_Record_Inquiry_Only Or 1 = 1 Then
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No ' FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End If
                End With
            End If
        Else
            grdARTCCTR2.Parent = splDetails.Panel1 ' tabCC.Tabs("Unmatched CC Transactions").TabPage
            grdARTCCTR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdARTCCTR2.DisplayLayout.GroupByBox.Hidden = False
            Show_Filter(grdARTCCTR2, True)
            Setup_grdARTCCTR2(True)
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() _
            {"ARTCCTR1", "ARTCCTR2", _
             "ARTPYMT1", "ARTPYMT2", "ARTPYMT3", "ARTPYMT4", "ARTPYMT5", _
             "ARTOPEN1", "ARTOPENX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Absx1.txtFor("BANK_CODE").Text = "" Then
            Absx1.txtFor("BANK_CODE").Text = ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE_CC") & ""
        End If

        Absx1.txtFor("PYMT_BATCH_NO").Text = ""
        Absx1.dteFor("PYMT_BATCH_DATE").Value = DBNull.Value
        Absx1.numFor("DEPOSIT_AMT").Value = 0
        Absx1.numFor("CC_FEE").Value = 0

        NO_UPDATE = False

        grdARTCCTR2.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        grdARTCCTRX.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()

        Load_ARTCCTRX()

        splDetails.Panel2Collapsed = True

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            HFs("CCTB_NO") = ASCMAIN1.Next_Control_No("ARTCCTR1.CCTB_NO")
        End If

        NO_UPDATE = False

        rowARTCCTR1 = Fill_Record("ARTCCTR1", HFs("CCTB_NO"), EntryMode = "N")
        If EntryMode = "N" Then
            rowARTCCTR1.Item("BANK_CODE") = HFs("BANK_CODE")
            If HFs("CCTB_DATE") & "" <> "" Then
                rowARTCCTR1.Item("CCTB_DATE") = HFs("CCTB_DATE")
            End If

            rowARTCCTR1.Item("STATUS") = "0"
            rowARTCCTR1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowARTCCTR1.Item("FILENAME") = FILENAME
        Else
            If rowARTCCTR1.Item("STATUS") & "" <> "0" Then
                This_Record_Inquiry_Only = True
            Else
                This_Record_Inquiry_Only = False
            End If
        End If
        Fill_Records("ARTCCTR2", HFs("CCTB_NO"))

        Sort_grdColumns(grdARTCCTR2, "CCTB_LNO")

        If EntryMode = "N" Then
            Try
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading CC Transactions")

                If isProcessingPayPal Then
                    ProcessPayPalData()
                Else
                    Dim r As Integer = 0
                    Dim r_start As Integer = 0
                    Dim CCTB_LNO As Integer = 0
                    For Each row As DataRow In dt.Rows
                        r += 1
                        ASCMAIN1.Progress("-", r & "/" & dt.Rows.Count)
                        If r = 8 Then
                            rowARTCCTR1.Item("CCTB_DATE") = row.Item(1)
                        ElseIf r = 9 Then
                            Dim CCTB_MEMO As String = row.Item(0) & row.Item(1) & String.Empty
                            CCTB_MEMO = CCTB_MEMO.Trim
                            If CCTB_MEMO.Length > rowARTCCTR1.Table.Columns("CCTB_MEMO").MaxLength Then
                                CCTB_MEMO = CCTB_MEMO.Substring(0, rowARTCCTR1.Table.Columns("CCTB_MEMO").MaxLength).Trim
                            End If
                            rowARTCCTR1.Item("CCTB_MEMO") = CCTB_MEMO
                        ElseIf r > 10 And r_start = 0 Then
                            If row.Item(0) & "" = "MID" And row.Item(1) & "" = "CardName" Then
                                r_start = r
                            End If
                        ElseIf r > r_start And r_start <> 0 Then
                            Dim invalidRecord As Boolean = True
                            For iloop As Integer = 1 To 5
                                If row.Item(iloop) & String.Empty <> String.Empty Then
                                    invalidRecord = False
                                End If
                            Next

                            If invalidRecord Then
                                Continue For
                            End If

                            Dim rowARTCCTR2 As DataRow = dst.Tables("ARTCCTR2").NewRow
                            rowARTCCTR2.Item("CCTB_NO") = HFs("CCTB_NO")
                            CCTB_LNO += 1
                            rowARTCCTR2.Item("CCTB_LNO") = CCTB_LNO

                            rowARTCCTR2.Item("CARD_NAME") = row.Item(1)
                            rowARTCCTR2.Item("TRANS_NAME") = row.Item(2)
                            rowARTCCTR2.Item("TIME") = row.Item(3)
                            rowARTCCTR2.Item("CARD_NUM") = row.Item(4)
                            rowARTCCTR2.Item("ENTRY_MODE") = row.Item(5)
                            rowARTCCTR2.Item("AUTH_CODE") = row.Item(6)
                            rowARTCCTR2.Item("TRANS_AMT") = Val((row.Item(7) & String.Empty).ToString.Replace(",", "")) ' row.Item(7)
                            rowARTCCTR2.Item("SERVICE_FEE") = 0
                            rowARTCCTR2.Item("TRANS_AMT_NET") = rowARTCCTR2.Item("TRANS_AMT")

                            rowARTCCTR2.Item("TRANSACTION") = row.Item(9)

                            rowARTCCTR2.Item("VOID") = row.Item(10)
                            rowARTCCTR2.Item("REFERENCE_NUMBER") = row.Item(11)
                            rowARTCCTR2.Item("EXPIRATION_DATE") = row.Item(12)
                            rowARTCCTR2.Item("AUTH_SOURCE") = row.Item(13)
                            rowARTCCTR2.Item("CARD_METHOD") = DBNull.Value
                            rowARTCCTR2.Item("POS_CAPABILITY") = row.Item(14)
                            rowARTCCTR2.Item("AVS") = row.Item(15)
                            rowARTCCTR2.Item("CVV2") = row.Item(16)
                            rowARTCCTR2.Item("TIP_AMT") = Val((row.Item(17) & String.Empty).ToString.Replace(",", "")) ' row.Item(17)
                            rowARTCCTR2.Item("ORDER_NUMBER") = row.Item(8)
                            rowARTCCTR2.Item("RETURN_COUNT") = row.Item(18)

                            ' If the Order Number is missing then see if we are to try a best match
                            Dim ORDR_NO As String = row.Item(8) & String.Empty
                            ORDR_NO = ORDR_NO.Trim
                            If ASCMAIN1.Running_in_VS AndAlso ORDR_NO = "000002700" Then Stop
                            If chkBestMatch.Checked And ORDR_NO.Length = 0 Then
                                Dim CUST_CREDIT_CARD_LAST4 As String = row.Item(4) & String.Empty
                                CUST_CREDIT_CARD_LAST4 = CUST_CREDIT_CARD_LAST4.Trim

                                Dim CCPA_DATE_AUTH As String = row.Item(3)
                                If Not IsDate(CCPA_DATE_AUTH) Then
                                    CCPA_DATE_AUTH = String.Empty
                                End If

                                Dim CUST_CREDIT_CARD_TYPE As String = String.Empty
                                Select Case (rowARTCCTR2.Item("CARD_NAME") & String.Empty).ToString.ToUpper
                                    Case "Visa".ToUpper
                                        CUST_CREDIT_CARD_TYPE = "VISA"
                                    Case "Discover".ToUpper
                                        CUST_CREDIT_CARD_TYPE = "DISC"
                                    Case "Master Card".ToUpper
                                        CUST_CREDIT_CARD_TYPE = "MSTR"
                                    Case "American Express".ToUpper
                                        CUST_CREDIT_CARD_TYPE = "AMEX"
                                End Select

                                Dim CCPA_AMT As Decimal = Val(row.Item(7))
                                Dim tblARTCCPA1 As DataTable = Nothing

                                If CUST_CREDIT_CARD_LAST4.Length > 0 _
                                    AndAlso CCPA_DATE_AUTH.Length > 0 _
                                    AndAlso CCPA_AMT > 0 Then

                                    CUST_CREDIT_CARD_LAST4 = CUST_CREDIT_CARD_LAST4.PadLeft(4, "0")
                                    ' The date in ARTCCPA1 is the date of the Authorization
                                    ' The date in the file is the date the sales order was charged.
                                    ' We assume all Consumer orders will ship within 2 weeks.
                                    ASCMAIN1.sql = "SELECT * FROM ARTCCPA1 " _
                                        & " WHERE CUST_CODE = 'CONSUMER'" _
                                        & " AND CUST_CREDIT_CARD_LAST4 = '" & CUST_CREDIT_CARD_LAST4 & "'" _
                                        & " And CCPA_AMT = " & CCPA_AMT _
                                        & " AND CCPA_DATE_AUTH BETWEEN '" & DateAdd(DateInterval.Day, numCCDays, CDate(CCPA_DATE_AUTH)).ToString("dd-MMM-yyyy") & "' AND '" & CDate(CCPA_DATE_AUTH).ToString("dd-MMM-yyyy") & "'" _
                                        & " AND CUST_CREDIT_CARD_TYPE = '" & CUST_CREDIT_CARD_TYPE & "'"

                                    tblARTCCPA1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                                    If tblARTCCPA1.Rows.Count = 1 Then
                                        ORDR_NO = tblARTCCPA1.Rows(0).Item("ORDR_NO")
                                    Else
                                        ' It may be the case the order short shipped or the prices changed
                                        ' Therefore drop the $$ value, grab the order No and look at the invoice total amount
                                        ASCMAIN1.sql = "SELECT * FROM ARTCCPA1 " _
                                            & " WHERE CUST_CODE = 'CONSUMER'" _
                                            & " AND CUST_CREDIT_CARD_LAST4 = '" & CUST_CREDIT_CARD_LAST4 & "'" _
                                            & " AND CCPA_DATE_AUTH BETWEEN '" & DateAdd(DateInterval.Day, numCCDays, CDate(CCPA_DATE_AUTH)).ToString("dd-MMM-yyyy") & "' AND '" & CDate(CCPA_DATE_AUTH).ToString("dd-MMM-yyyy") & "'" _
                                            & " AND CUST_CREDIT_CARD_TYPE = '" & CUST_CREDIT_CARD_TYPE & "'"

                                        tblARTCCPA1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                                        If tblARTCCPA1.Rows.Count = 1 Then
                                            ORDR_NO = tblARTCCPA1.Rows(0).Item("ORDR_NO")
                                            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE ORDR_NO = '" & ORDR_NO & "' and INV_TOTAL_AMOUNT = " & CCPA_AMT)
                                            ORDR_NO = String.Empty
                                            If rowSOTINVH1 IsNot Nothing Then
                                                ORDR_NO = tblARTCCPA1.Rows(0).Item("ORDR_NO")
                                            End If
                                        ElseIf tblARTCCPA1.Rows.Count > 1 Then
                                            ASCMAIN1.sql = "SELECT * FROM SOTINVH1 " _
                                                & " WHERE ORDR_NO IN ( " _
                                                & " SELECT ORDR_NO FROM ARTCCPA1 " _
                                                & " WHERE CUST_CODE = 'CONSUMER' " _
                                                & " AND CCPA_DATE_AUTH BETWEEN '" & DateAdd(DateInterval.Day, numCCDays, CDate(CCPA_DATE_AUTH)).ToString("dd-MMM-yyyy") & "' AND '" & CDate(CCPA_DATE_AUTH).ToString("dd-MMM-yyyy") & "'" _
                                                & " AND CUST_CREDIT_CARD_LAST4 = '" & CUST_CREDIT_CARD_LAST4 & "'" _
                                                & " AND CUST_CREDIT_CARD_TYPE = '" & CUST_CREDIT_CARD_TYPE & "'" _
                                                & " )" _
                                                & " AND INV_TOTAL_AMOUNT = " & CCPA_AMT
                                            Dim tblSOTINVH1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                                            If tblSOTINVH1.Rows.Count = 1 Then
                                                ORDR_NO = tblSOTINVH1.Rows(0).Item("ORDR_NO")
                                            End If
                                        End If
                                    End If

                                    If ORDR_NO.Length > 0 Then
                                        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = '" & ORDR_NO & "'")
                                        ORDR_NO = String.Empty
                                        If rowSOTORDR1 IsNot Nothing Then
                                            ORDR_NO = rowSOTORDR1.Item("ORDR_NO_WEB") & String.Empty
                                        End If
                                    End If

                                    If ORDR_NO.Length > 0 Then
                                        ASCMAIN1.sql = "Select ARTCCTR2.* from ARTCCTR1, ARTCCTR2 " _
                                            & " where ARTCCTR1.CCTB_NO = ARTCCTR2.CCTB_NO" _
                                            & " AND ARTCCTR2.ORDER_NUMBER = '" & ORDR_NO & "' and ARTCCTR2.TRANS_AMT > 0" _
                                            & " AND ARTCCTR1.STATUS <> 'D'"
                                        Dim rowARTCCTR2x As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                                        If rowARTCCTR2x Is Nothing AndAlso dst.Tables("ARTCCTR2").Select("ORDER_NUMBER = '" & ORDR_NO & "' And TRANS_AMT > 0").Length = 0 Then
                                            rowARTCCTR2.Item("ORDER_NUMBER") = ORDR_NO
                                        End If
                                    End If
                                ElseIf CUST_CREDIT_CARD_LAST4.Length > 0 _
                                    AndAlso CCPA_DATE_AUTH.Length > 0 _
                                    AndAlso CCPA_AMT < 0 Then

                                    CUST_CREDIT_CARD_LAST4 = CUST_CREDIT_CARD_LAST4.PadLeft(4, "0")
                                    ASCMAIN1.sql = "select DISTINCT ORDR_CUST_PO from sotinvh1 " _
                                        & " where INV_TYPE = 'C'" _
                                        & " AND cust_code = 'CONSUMER'" _
                                        & " AND ORDR_CUST_PO IS NOT NULL" _
                                        & " AND INV_TOTAL_AMOUNT = " & CCPA_AMT _
                                        & " and ORDR_CUST_PO IN" _
                                        & " (" _
                                        & " SELECT ORDR_NO_WEB FROM SOTORDR1" _
                                        & " WHERE ORDR_NO IN" _
                                        & " ( " _
                                        & " SELECT ORDR_NO FROM ARTCCPA1 " _
                                        & " WHERE CUST_CODE = 'CONSUMER' " _
                                        & " AND CUST_CREDIT_CARD_LAST4 = '" & CUST_CREDIT_CARD_LAST4 & "' " _
                                        & " AND CCPA_DATE_AUTH BETWEEN '" & DateAdd(DateInterval.Day, -60, CDate(CCPA_DATE_AUTH)).ToString("dd-MMM-yyyy") & "' AND '" & CDate(CCPA_DATE_AUTH).ToString("dd-MMM-yyyy") & "'" _
                                        & " AND CUST_CREDIT_CARD_TYPE = '" & CUST_CREDIT_CARD_TYPE & "'" _
                                        & " )" _
                                        & " )"

                                    Dim tblSOTINVH1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                                    If tblSOTINVH1.Rows.Count = 1 Then
                                        ORDR_NO = tblSOTINVH1.Rows(0).Item("ORDR_CUST_PO") & String.Empty
                                        ORDR_NO = ORDR_NO.Replace("'", "")
                                    End If

                                    If ORDR_NO.Length > 0 Then
                                        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO_WEB = '" & ORDR_NO & "'")
                                        ORDR_NO = String.Empty
                                        If rowSOTORDR1 IsNot Nothing Then
                                            ORDR_NO = rowSOTORDR1.Item("ORDR_NO_WEB") & String.Empty
                                        End If
                                    End If

                                    If ORDR_NO.Length > 0 Then
                                        rowARTCCTR2.Item("ORDER_NUMBER") = ORDR_NO
                                    End If
                                End If
                            End If

                            For I As Integer = 0 To 18
                                'rowARTCCTR2.Item(I + 2) = row.Item(I)

                                With dst.Tables("ARTCCTR2").Columns(I + 2)
                                    If .DataType = GetType(System.String) Then
                                        If Len(rowARTCCTR2.Item(I + 2) & "") > .MaxLength Then
                                            rowARTCCTR2.Item(I + 2) = Mid((rowARTCCTR2.Item(I + 2) & ""), 1, .MaxLength)
                                        End If
                                    End If
                                End With
                            Next

                            Dim skip_record As Boolean = False
                            If rowARTCCTR2.Item("CARD_NAME") & "" = "" Then
                                If rowARTCCTR2.Item("TRANS_NAME") & "" = "" And _
                                    rowARTCCTR2.Item("CARD_NUM") & "" = "" And _
                                    rowARTCCTR2.Item("AUTH_CODE") & "" = "" And _
                                    rowARTCCTR2.Item("EXPIRATION_DATE") & "" = "" And _
                                    rowARTCCTR2.Item("TRANS_AMT") & "" = "" Then
                                    skip_record = True
                                End If
                            End If

                            If skip_record Then
                            Else
                                dst.Tables("ARTCCTR2").Rows.Add(rowARTCCTR2)

                                ' AR Item Matching Process
                                Dim TRANS_AMT As Decimal = Val(rowARTCCTR2.Item("TRANS_AMT") & "")
                                Match_AR_Item(TRANS_AMT, rowARTCCTR2)

                            End If
                        End If
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Process This File")
                NO_UPDATE = True

            Finally
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            End Try
        Else
            dst.AcceptChanges()
        End If

        'For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("TRANS_AMT > 0 and INV_TYPE = 'C'")
        '    rowARTCCTR2.Item("INVAILD_TRANS_AMT") = "1"
        'Next

        'For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("TRANS_AMT < 0 and INV_TYPE = 'I'")
        '    rowARTCCTR2.Item("INVAILD_TRANS_AMT") = "1"
        'Next

        grdARTCCTR2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        Display_Totals()
    End Sub

    Sub Delete_Record()
        BeginTrans()
        ASCDATA1.ExecuteSQL("Update ARTCCTR1 Set STATUS = 'D' where CCTB_NO = '" & HFs("CCTB_NO") & "'")
        CommitTrans("Record marked as Deleted")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Display_Totals()

        INIT_LAST("ARTCCTR1", , , True)
        INIT_LAST("ARTCCTR2", , , True)

        Update_Record_TDA("ARTCCTR1")
        Update_Record_TDA("ARTCCTR2")

        CommitTrans("Update Complete")

        If EntryMode = "N" Then
            Try
                My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Archive") & "Cynergy\" & ASCMAIN1.DBS_COMPANY & "\" & HFs("CCTB_NO") & "\")
                Dim F As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                My.Computer.FileSystem.MoveFile(FILENAME, ASCMAIN1.Folders("Archive") & "Cynergy\" & ASCMAIN1.DBS_COMPANY & "\" & HFs("CCTB_NO") & "\" & F.Name)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Occurred trying to Archive File")
            End Try
        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CCTB_NO"
                'sql_where = "STATUS = '0'"
        End Select
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCCTRX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTCCTR2, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Select Selected", "Select All", "De-Select All", "Select AR Item", "Clear AR Item from Selected Rows", "Refresh AR Item Matches for Selected Rows", "Sales Order Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.SourceControl.Name
            Case "grdARTCCTR2"
                tlb.Tools("Select All").SharedProps.Visible = (EntryMode = "S")
                tlb.Tools("De-Select All").SharedProps.Visible = (EntryMode = "S")
                tlb.Tools("Select Selected").SharedProps.Visible = (EntryMode = "S")
                tlb.Tools("Select AR Item").SharedProps.Visible = (EntryMode = "S")
                tlb.Tools("Clear AR Item from Selected Rows").SharedProps.Visible = (EntryMode = "S")
                tlb.Tools("Refresh AR Item Matches for Selected Rows").SharedProps.Visible = (EntryMode = "S")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdARTPYMT3"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Sales Order Inquiry"
                If grd.ActiveRow Is Nothing Then
                    Exit Sub
                End If

                Dim ORDR_CUST_PO As String = grd.ActiveRow.Cells("ORDER_NUMBER").Value & String.Empty
                ORDR_CUST_PO = ORDR_CUST_PO.Trim

                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty
                CUST_CODE = CUST_CODE.Trim

                If ORDR_CUST_PO.Length = 0 OrElse CUST_CODE.Length = 0 Then
                    MessageBox.Show("Missing Web Order No or Customer Code", "Sales Order Inquiry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_CUST_PO = :PARM1 and CUST_CODE = :PARM2", "VV", New Object() {ORDR_CUST_PO, CUST_CODE})
                If rowSOTORDR1 Is Nothing Then
                    MessageBox.Show("Could not locate Sales Order associated with Web Order Number.", "Sales Order Inquiry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Context_Launch("View", rowSOTORDR1.Item("ORDR_NO"), e.Tool.Key, "SOFORDRI")

            Case "Select Selected"
                ASCMAIN1.Progress("Now Selecting Selected Rows")
                For Each grow As UltraWinGrid.UltraGridRow In grdARTCCTR2.Selected.Rows
                    ASCMAIN1.Progress("-", grow.Cells("INV_NUM").Value & "-" & grow.Cells("ORDER_NUMBER").Value & "-" & grow.Cells("REFERENCE_NUMBER").Value)
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next
                Display_Totals()
                ASCMAIN1.Progress("")

            Case "Select All", "De-Select All"
                For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("")
                    rowARTCCTR2.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
                Display_Totals()

            Case "Clear AR Item from Selected Rows", "Refresh AR Item Matches for Selected Rows"
                If grdARTCCTR2.Selected.Rows.Count = 0 Then
                    If grdARTCCTR2.ActiveRow IsNot Nothing Then
                        grdARTCCTR2.ActiveRow.Selected = True
                    End If
                End If

                If grdARTCCTR2.Selected.Rows.Count <> 0 Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdARTCCTR2.Selected.Rows
                        If e.Tool.Key = "Clear AR Item from Selected Rows" Then
                            grow.Cells("CUST_CODE").Value = DBNull.Value
                            grow.Cells("INV_TYPE").Value = DBNull.Value
                            grow.Cells("INV_NUM").Value = DBNull.Value
                            grow.Cells("INV_TOTAL_AMOUNT").Value = DBNull.Value
                            grow.Cells("INV_BALANCE").Value = DBNull.Value
                            grow.Update()
                        Else
                            Dim TRANS_AMT As Decimal = Val(grow.Cells("TRANS_AMT").Value & "")
                            Dim rowARTCCTR2 As DataRow = dst.Tables("ARTCCTR2").Rows.Find _
                                                         (New Object() {grow.Cells("CCTB_NO").Value, _
                                                                        grow.Cells("CCTB_LNO").Value})
                            Match_AR_Item(TRANS_AMT, rowARTCCTR2)
                        End If
                    Next
                End If
                Display_Totals()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCTR2" Then

            Select Case e.Tool.Key
                Case "Select AR Item"
                    Dim sql_where As String = ""
                    ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("INV_NUM", "ARTOPEN1", sql_where)
                    If ASCMAIN1.CodeSelector.SQL <> "" Then
                        ASCMAIN1.CodeSelector.MultipleSelections = False
                        Dim F As New ASFCODE1
                        F.ShowDialog()
                        F.Dispose()
                        If ASCMAIN1.CodeSelector.Selections <> 0 Then
                            grdARTCCTR2.ActiveRow.Cells("CUST_CODE").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CUST_CODE")
                            grdARTCCTR2.ActiveRow.Cells("INV_TYPE").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_TYPE")
                            grdARTCCTR2.ActiveRow.Cells("INV_NUM").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_NUM")
                            grdARTCCTR2.ActiveRow.Cells("INV_TOTAL_AMOUNT").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_TOTAL_AMOUNT")
                            grdARTCCTR2.ActiveRow.Cells("INV_BALANCE").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_BALANCE")
                        End If
                    End If
            End Select
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("New", e)
                End If
            Case "CCTB_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CCTB_NO"
                Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        If ScreenMode Then
            Display_Totals()
        End If
    End Sub
#End Region


#Region "grdARTCCTR2"

    Private Sub grdARTCCTR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTCCTR2.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdARTCCTR2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTCCTR2.BeforeCellUpdate
        'Select Case e.Cell.Column.Key
        '    Case "CUST_PYMT_REF_DATE"
        '        If "" = "" Then
        '            Dim currYear As String = Mid(ASCMAIN1.CYP, 1, 4)
        '            'grdARTCCTR2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
        '        End If
        'End Select
    End Sub

    Private Sub grdARTCCTR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCCTR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "CUST_CODE"
            '    If e.Cell.Text = "" Then

            '    Else

            '        grdCodeDesc(grdARTCCTR2, "ARTCUST1", "CUST_CODE", "CUST_NAME")
            '        If grdARTCCTR2.ActiveRow.Cells("CUST_NAME").Text = "" Then
            '            grdARTCCTR2.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
            '        Else
            '            grdARTCCTR2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
            '            grdARTCCTR2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
            '        End If
            '    End If
        End Select
    End Sub

    Private Sub grdARTCCTR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCCTR2.AfterRowActivate

        If EntryMode = "S" AndAlso grdARTCCTR2.ActiveRow.Cells("MISSING_ORDER_NUMBER").Value & String.Empty = "1" Then
            Me.Cursor = Cursors.WaitCursor
            Dim CARD_NUM As String = grdARTCCTR2.ActiveRow.Cells("CARD_NUM").Value & String.Empty
            Dim TRANS_AMT As Decimal = Val(grdARTCCTR2.ActiveRow.Cells("TRANS_AMT").Value & String.Empty)
            Dim CARD_NAME As String = grdARTCCTR2.ActiveRow.Cells("CARD_NAME").Value & String.Empty
            Dim TIME As String = grdARTCCTR2.ActiveRow.Cells("TIME").Value & String.Empty

            Select Case (CARD_NAME & String.Empty).ToString.ToUpper
                Case "Visa".ToUpper
                    CARD_NAME = "VISA"
                Case "Discover".ToUpper
                    CARD_NAME = "DISC"
                Case "Master Card".ToUpper
                    CARD_NAME = "MSTR"
                Case "American Express".ToUpper
                    CARD_NAME = "AMEX"
            End Select

            ASCMAIN1.sql = "Select * from ARTOPEN1 WHERE (INV_TYPE, INV_NUM) IN " _
                & " ( " _
                & " SELECT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO" _
                & " FROM ARTCCPA1, SOTINVH1 " _
                & " WHERE ARTCCPA1.CUST_CODE = 'CONSUMER'" _
                & " AND ARTCCPA1.CUST_CREDIT_CARD_LAST4 LIKE '%" & CARD_NUM & "'" _
                & " AND ARTCCPA1.CCPA_DATE_AUTH BETWEEN '" & DateAdd(DateInterval.Day, numCCDays, CDate(TIME)).ToString("dd-MMM-yyyy") & "' AND '" & CDate(TIME).ToString("dd-MMM-yyyy") & "'" _
                & " AND ARTCCPA1.CUST_CREDIT_CARD_TYPE = '" & CARD_NAME & "'" _
                & " AND SOTINVH1.ORDR_NO = ARTCCPA1.ORDR_NO" _
                & ")"

            Fill_Records("ARTOPENX", String.Empty, True, ASCMAIN1.sql)
            splDetails.Panel2Collapsed = False
            splDetails.Panel2MinSize = 300
            Me.Cursor = Cursors.Default
        Else
            splDetails.Panel2Collapsed = True
        End If
    End Sub

    Private Sub grdARTCCTR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCCTR2.BeforeRowUpdate
        With grdARTCCTR2
            'If e.Row.Cells("CUST_CODE").Text = "" Then
            '    MsgBox("Missing Value for Customer Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'Else
            '    LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
            '    If cdr Is Nothing Then
            '        MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '        e.Cancel = True
            '    End If
            'End If

            If Not e.Cancel Then
                If e.Row.Cells("CCTB_NO").Text = "" Then
                    .ActiveRow.Cells("CCTB_NO").Value = Absx1.CtlFor("CCTB_NO").Text
                    .ActiveRow.Cells("CCTB_LNO").Value = Val(dst.Tables("ARTCCTR2").Compute("Max(CCTB_LNO)", "") & "") + 1
                    .ActiveRow.Cells("CCTB_STATUS").Value = "0"
                    'LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value)
                    'e.Row.Cells("CUST_NAME").Value = cdr.Item("CUST_NAME")
                End If
            End If

        End With
    End Sub

    Private Sub grdARTCCTR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCCTR2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTCCTR2, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTCCTR2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCCTR2.Error
        grdARTCCTR2.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdARTCCTR2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdARTCCTR2.KeyPress
        'If grdARTCCTR2.ActiveRow IsNot Nothing Then
        '    Try
        '        If grdARTCCTR2.ActiveCell.Column.Key = "CUST_NAME" Then
        '            If grdARTCCTR2.ActiveRow.Cells("CUST_CODE").Text <> "" Then
        '                e.KeyChar = Chr(0)
        '                e.Handled = True
        '            End If
        '        ElseIf grdARTCCTR2.ActiveCell.Column.Key = "CUST_PYMT_REF_DATE" Then

        '            If e.KeyChar = "" Then
        '                If "" = "" Then
        '                    Dim year As String = Mid(ASCMAIN1.CYP, 0, 4)
        '                    grdARTCCTR2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
        '                End If
        '            End If

        '        End If
        '    Catch ex As Exception

        'End Try
        'End If
    End Sub

    Private Sub grdARTCCTR2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTCCTR2.BeforeExitEditMode
        'With grdARTCCTR2.ActiveCell
        '    Select Case .Column.Key
        '        Case "CUST_CODE"
        '            If .Text <> "" Then
        '                .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
        '            End If

        '        Case "CUST_PYMT_REF_DATE"
        '            'Stop
        '            Dim DT As String = .EditorResolved.CurrentEditText
        '            If Len(DT) = 10 And Mid(DT, 7, 4) = "    " Then
        '                .Value = Mid(DT, 1, 6) & Now.Year
        '            End If
        '    End Select
        'End With
    End Sub

    Private Sub grdARTCCTR2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCCTR2.InitializeRow
        If (Val(e.Row.Cells("TRANS_AMT").Value & "") > 0 And e.Row.Cells("INV_TYPE").Value & "" = "C") _
        Or (Val(e.Row.Cells("TRANS_AMT").Value & "") < 0 And e.Row.Cells("INV_TYPE").Value & "" = "I") Then
            e.Row.Cells("INV_NUM").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("INV_NUM").ToolTipText = "CC Transaction Type on this record is not normally matched to this type of AR Item"
            e.Row.Cells("TRANS_AMT").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("INV_NUM").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("INV_NUM").ToolTipText = ""
            e.Row.Cells("TRANS_AMT").Appearance.ForeColor = Drawing.Color.Empty
        End If


    End Sub

#End Region

    Private Sub grdARTCCTRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCCTRX.DoubleClickRow
        If grdARTCCTRX.ActiveRow IsNot Nothing AndAlso grdARTCCTRX.ActiveRow.IsDataRow Then
            Absx1.txtFor("CCTB_NO").Text = grdARTCCTRX.ActiveRow.Cells("CCTB_NO").Text
            Click_Command("Edit")
        End If

    End Sub

    Sub Load_ARTCCTRX(Optional header_only As Boolean = False)
        If SELECTION_NO = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Payment Batch Data")

        Dim OPS_YYYYPP As String = ""

        Dim STATUS As String = "0"


        If chkUnappliedOnly.Checked Then
            STATUS = "0"
            grdARTCCTRX.Text = "Transactions Loaded, but not Applied"
        Else
            STATUS = "1"
            grdARTCCTRX.Text = "Transactions Loaded and Applied"
        End If
        'Fill_Records("ARTCCTRX")
        ASCMAIN1.sql = "SELECT ARTCCTR1.* from ARTCCTR1 where NVL(ARTCCTR1.STATUS,'0') = '" & STATUS & "'"
        Fill_Records("ARTCCTRX", , , ASCMAIN1.sql)
        Sort_grdColumns(grdARTCCTRX, "CCTB_NO".ToLower)

        If Not header_only Then
            ASCMAIN1.sql = "Select ARTCCTR2.* from ARTCCTR2,ARTCCTR1 where ARTCCTR2.PYMT_BATCH_NO is Null" & vbCrLf _
                & " and ARTCCTR1.CCTB_NO = ARTCCTR2.CCTB_NO and ARTCCTR1.STATUS = '" & "0" & "'"
            Fill_Records("ARTCCTR2", "", True, ASCMAIN1.sql)

            'For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("TRANS_AMT > 0 and INV_TYPE = 'C'")
            '    rowARTCCTR2.Item("INVAILD_TRANS_AMT") = "1"
            'Next

            'For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("TRANS_AMT < 0 and INV_TYPE = 'I'")
            '    rowARTCCTR2.Item("INVAILD_TRANS_AMT") = "1"
            'Next

            '  grdARTCCTR2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        End If

     
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdARTCCTRX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTCCTRX.InitializeLayout

    End Sub

    Sub Display_Totals()
        Static processing As Boolean = False
        If processing Then Exit Sub
        If EntryMode = "S" Then
            dst.Tables("ARTCCTRT").Rows.Clear()
            processing = True

            Dim TOTAL_CC As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(TRANS_AMT)", "SEL = '1'") & "")
            Dim TOTAL_AR As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(INV_BALANCE)", "SEL = '1'") & "")
            Dim CC_FEE As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(SERVICE_FEE)", "SEL = '1'") & "") * -1

            dst.Tables("ARTCCTRT").Rows.Add(New Object() {1, "CC Trans Amt", TOTAL_CC})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {2, "AR Balance", TOTAL_AR})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {3, "Non-AR", TOTAL_CC - TOTAL_AR})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {4})
            Absx1.numFor("CC_FEE").Value = CC_FEE

            dst.Tables("ARTCCTRT").Rows.Add(New Object() {5, "CC Fee", CC_FEE})
            Dim DEPOSIT_AMT As Decimal = Val(Absx1.numFor("DEPOSIT_AMT").Value & "")
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {6, "Deposit Amt", DEPOSIT_AMT})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {7, "Out of Balance", DEPOSIT_AMT - (TOTAL_CC - CC_FEE)})
            processing = False
        Else
            Dim TOTAL_TRX As Int64 = dst.Tables("ARTCCTR2").Select("").Length
            Dim TOTAL_SLS As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(TRANS_AMT)", "TRANS_NAME = 'Sale' OR TRANS_NAME = 'Prior auth sale' OR TRANS_NAME = 'Mail/Telephone'") & "")
            Dim TOTAL_RTN As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(TRANS_AMT)", "TRANS_NAME = 'Return'") & "")
            Dim TOTAL_UNMATCHED As Decimal = Val(dst.Tables("ARTCCTR1").Rows(0).Item("TOTAL_UNMATCHED") & String.Empty)

            rowARTCCTR1.Item("TOTAL_TRX") = TOTAL_TRX
            rowARTCCTR1.Item("TOTAL_SLS") = TOTAL_SLS
            rowARTCCTR1.Item("TOTAL_RTN") = TOTAL_RTN

            dst.Tables("ARTCCTRT").Rows.Clear()

            ' dst.Tables("ARTCCTRT").Rows.Add(New Object() {1, "Trans", TOTAL_TRX})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {2, "Sales", TOTAL_SLS})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {3, "Returns", TOTAL_RTN})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {4, "Net", TOTAL_SLS + TOTAL_RTN})
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {5})

            Dim KEY As Integer = 10
            Dim TRANS_AMT_TOTAL As Decimal = 0
            For Each row As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("ARTCCTR2"), New String() {"CARD_NAME"}).Rows
                Dim CARD_NAME As String = row.Item(0) & ""
                Dim TRANS_AMT As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(TRANS_AMT)", "CARD_NAME = '" & CARD_NAME & "'") & "")
                KEY += 1
                dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY, CARD_NAME, TRANS_AMT})
                TRANS_AMT_TOTAL += TRANS_AMT
            Next
            KEY += 1
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY, "All Cards", TRANS_AMT_TOTAL})

            KEY += 1
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY})

            Dim INV_BALANCE As Decimal = Val(dst.Tables("ARTCCTR2").Compute("SUM(INV_BALANCE)", "") & "")
            KEY += 1
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY, "Total AR", INV_BALANCE})
            rowARTCCTR1.Item("INV_BALANCE") = INV_BALANCE
            KEY += 1
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY, "Difference", INV_BALANCE - TRANS_AMT_TOTAL})
            KEY += 1
            dst.Tables("ARTCCTRT").Rows.Add(New Object() {KEY, "Not Comp", TOTAL_UNMATCHED})
        End If
    End Sub

    Private Sub grdARTCCTRT_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCCTRT.InitializeRow
        If e.Row.Cells("STATUS").Value & "" = "Total" _
        Or e.Row.Cells("STATUS").Value & "" = "Net" _
        Or e.Row.Cells("STATUS").Value & "" = "Deposit Total" _
        Or e.Row.Cells("STATUS").Value & "" = "All Cards" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        End If
        If e.Row.Cells("STATUS").Value & "" = "Difference" _
        Or e.Row.Cells("STATUS").Value & "" = "Out of Balance" Then
            If Val(e.Row.Cells("AMT").Value & "") <> 0 Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
            End If
        End If

    End Sub

    Private Sub grdARTCCTRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCCTRX.InitializeRow
        If Val(e.Row.Cells("TOTAL_SLS").Value & "") + Val(e.Row.Cells("TOTAL_RTN").Value & "") <> Val(e.Row.Cells("INV_BALANCE").Value & "") Then
            e.Row.Cells("INV_BALANCE").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Sub Setup_grdARTCCTR2(hidden As Boolean)
        For Each COLUMN_NAME As String In New String() {"CCTB_NO", "SEL", "INV_BALANCE_SEL", "TRANS_AMT_SEL"}
            grdARTCCTR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = hidden
        Next
    End Sub

    Private Sub tabCC_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabCC.SelectedTabChanged
        Setup_tabCC()
    End Sub

    Sub Setup_tabCC()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Screen Control").Visible = (tabCC.SelectedTab.Key = "Batches")
        UltraExplorerBar1.Groups("Credit Card Deposit").Visible = (tabCC.SelectedTab.Key = "Unmatched CC Transactions")
        grpCCTrans.Visible = (tabCC.SelectedTab.Key = "Batches")
        grpDeposit.Visible = (tabCC.SelectedTab.Key = "Unmatched CC Transactions") And ScreenMode
        If (tabCC.SelectedTab.Key = "Unmatched CC Transactions") And ScreenMode Then
            tabCC.Tabs("Batches").Enabled = False
        Else
            tabCC.Tabs("Batches").Enabled = True
        End If

        chkBestMatch.Visible = Not ScreenMode

    End Sub

    Sub Update_Deposit()

        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        Absx1.txtFor("PYMT_BATCH_NO").Text = PYMT_BATCH_NO
        With rowARTPYMT1
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_DATE") = Absx1.dteFor("PYMT_BATCH_DATE").Value
            .Item("BANK_CODE") = Absx1.txtFor("BANK_CODE").Text
            .Item("STATUS") = "1"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("PYMT_APPL_ONLY") = DBNull.Value
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1
            .Item("PYMT_SOURCE") = "CC"
        End With
        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        Dim CUST_CODE As String = ""
        Dim CUST_NAME As String = ""
        Dim rowARTCUST1 As DataRow = Nothing

        dst.Tables("ARTCCTR2").AcceptChanges()

        Dim CC_FEE As Decimal = Val(Absx1.numFor("CC_FEE").Value & "")
        Dim TOTAL_CC As Decimal = 0
        Dim TOTAL_AR As Decimal = 0

        Dim PYMT_BATCH_LNO As Int32 = 0
        For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("SEL='1'")
            If CUST_CODE <> rowARTCCTR2.Item("CUST_CODE") & "" Then
                CUST_CODE = rowARTCCTR2.Item("CUST_CODE") & ""
                If CUST_CODE = "" Then
                    CUST_NAME = ""
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    CUST_NAME = rowARTCUST1.Item("CUST_NAME") & ""
                End If
            End If

            If CUST_CODE = "" Then
                CUST_CODE = "CONSUMER"
                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                CUST_NAME = rowARTCUST1.Item("CUST_NAME") & ""
            End If

            Dim TRANS_AMT As Decimal = Val(rowARTCCTR2.Item("TRANS_AMT") & "")


            Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
            With rowARTPYMT2
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                PYMT_BATCH_LNO += 1
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_NAME") = CUST_NAME
                .Item("CUST_PYMT_REF_NO") = ""
                .Item("CUST_PYMT_REF_DATE") = DATETIME_STAMP.Date
                .Item("CUST_PYMT_AMT") = TRANS_AMT
                .Item("PYMT_STATUS") = "2"
                .Item("CUST_PYMT_AMT_CURR") = TRANS_AMT
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("CUST_PYMT_CC_NO") = rowARTCCTR2.Item("CARD_NUM")
                .Item("CUST_PYMT_CC_TRANS_ID") = rowARTCCTR2.Item("AUTH_CODE")
                .Item("CUST_PYMT_CC_TRANS_DATE") = rowARTCCTR2.Item("TIME")

                .Item("CURR_CODE") = "USD"
                .Item("CURR_EXCH_RATE") = 1

                ' paypal does not provided the Credit Card Exp Date
                If IsDate(rowARTCCTR2.Item("EXPIRATION_DATE") & String.Empty) Then
                    .Item("CUST_PYMT_CC_EXP_MMYY") = Format(rowARTCCTR2.Item("EXPIRATION_DATE"), "MMyy")
                End If

                .Item("PYMT_NOTE") = rowARTCCTR2.Item("ORDER_NUMBER")
                Dim CARD_NAME = rowARTCCTR2.Item("CARD_NAME") & ""
                Dim CUST_CREDIT_CARD_TYPE As String = Mid(CARD_NAME, 1, 4)
                Select Case CARD_NAME
                    Case "American Express"
                        CUST_CREDIT_CARD_TYPE = "AMEX"
                    Case "Visa"
                        CUST_CREDIT_CARD_TYPE = "VISA"
                    Case "Master Card"
                        CUST_CREDIT_CARD_TYPE = "MSTR"
                    Case "Discover"
                        CUST_CREDIT_CARD_TYPE = "DISC"
                End Select
                .Item("CUST_CREDIT_CARD_TYPE") = CUST_CREDIT_CARD_TYPE
            End With
            dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

            TOTAL_CC += TRANS_AMT

            rowARTCCTR2.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            rowARTCCTR2.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO

            If TRANS_AMT <> 0 Then
                Dim rowARTOPEN1 As DataRow = Nothing
                If CUST_CODE <> "" AndAlso rowARTCCTR2.Item("INV_NUM") & "" <> "" Then
                    rowARTOPEN1 = dst.Tables("ARTOPEN1").Rows.Find _
                                                 (New String() {rowARTCCTR2.Item("CUST_CODE"), _
                                                                rowARTCCTR2.Item("INV_TYPE"), _
                                                                rowARTCCTR2.Item("INV_NUM")})

                    If rowARTOPEN1 Is Nothing Then
                        rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {rowARTCCTR2.Item("CUST_CODE"), _
                                                                rowARTCCTR2.Item("INV_TYPE"), _
                                                                rowARTCCTR2.Item("INV_NUM")}, False, False)
                    End If

                    If rowARTOPEN1 Is Nothing Then
                        MsgBox("Cannot find AR Item " & rowARTCCTR2.Item("INV_TYPE") & "-" & rowARTCCTR2.Item("INV_NUM") _
                               & vbCrLf & " in the Open AR Items File for Customer " & rowARTCCTR2.Item("CUST_CODE"), _
                                MsgBoxStyle.OkOnly, "Update Cancelled")
                        Exit Sub
                    End If
                End If
                If rowARTOPEN1 Is Nothing Then
                    If "" = "write off cash receipt" Then
                        Dim rowARTPYMT4 As DataRow = dst.Tables("ARTPYMT4").NewRow
                        With rowARTPYMT4
                            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                            .Item("PYMT_BATCH_GLNO") = 1
                            .Item("GL_DIST_AMT") = -1 * TRANS_AMT
                            .Item("GL_DIST_REF") = DBNull.Value
                            .Item("GL_DIST_AUTO") = "1"
                            .Item("ACCT_CODE") = ROWs("ARTPARM1").Item("AR_PARM_ACCT_CODE_CC_NON_AR")
                            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                            .Item("GL_DIST_AMT_CURR") = -1 * TRANS_AMT
                        End With
                        dst.Tables("ARTPYMT4").Rows.Add(rowARTPYMT4)
                    Else
                        rowARTPYMT2.Item("PYMT_STATUS") = "1"
                        rowARTPYMT2.Item("CUST_PYMT_REF_NO") = rowARTCCTR2.Item("ORDER_NUMBER")
                    End If
                Else
                    Dim INV_BALANCE As Decimal = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
                    Dim INV_BALANCE_EXPECTED As Decimal = Val(rowARTCCTR2.Item("INV_BALANCE") & "")
                    If INV_BALANCE <> INV_BALANCE_EXPECTED Then
                        MsgBox("Problem with Invoice " & rowARTOPEN1.Item("INV_NUM") _
                               & vbCrLf & "Invoice Balance (" & Format(INV_BALANCE, "#.00") _
                               & ") is not the expected value (" & Format(INV_BALANCE_EXPECTED, "#.00") & ")", _
                                MsgBoxStyle.OkOnly, "Update Cancelled")
                        Exit Sub
                    End If
                    If INV_BALANCE <> 0 Then
                        'If System.Math.Sign(INV_BALANCE) <> System.Math.Sign(TRANS_AMT) Then
                        '    MsgBox("Problem with Invoice " & rowARTOPEN1.Item("INV_NUM") _
                        '           & vbCrLf & "Invoice Balance (" & Format(INV_BALANCE, "#.00") _
                        '           & ") is not the same sign as the CC Trans Amount (" & Format(TRANS_AMT, "#.00") & ")", _
                        '            MsgBoxStyle.OkOnly, "Update Cancelled")
                        '    Exit Sub
                        'End If
                    End If

                    Dim INV_PMT As Decimal = 0
                    Dim INV_WRITE_OFF As Decimal = 0

                    Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
                    With rowARTPYMT3
                        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                        .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                        .Item("PYMT_BATCH_ILNO") = 1
                        For Each C As String In New String() _
                            {"INV_TYPE", "INV_NUM", "INV_DATE", "INV_DUE_DATE", _
                             "REASON_CODE", "CUST_CODE_SO", "CUST_STORE_NO", _
                             "INV_CUST_PO", "INV_BALANCE", _
                             "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}
                            .Item(C) = rowARTOPEN1.Item(C)
                        Next

                        For Each C As String In New String() {"INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF"}
                            .Item(C) = 0
                        Next

                        'INV_PMT = INV_BALANCE
                        'If System.Math.Abs(INV_BALANCE) > System.Math.Abs(TRANS_AMT) Then
                        '    INV_PMT = TRANS_AMT
                        'End If
                        'INV_WRITE_OFF = INV_BALANCE - INV_PMT

                        INV_PMT = TRANS_AMT
                        INV_WRITE_OFF = 0

                        .Item("INV_PMT") = INV_PMT
                        .Item("INV_WRITE_OFF") = INV_WRITE_OFF
                        .Item("INV_BALANCE_NEW") = INV_BALANCE - INV_PMT - INV_WRITE_OFF

                        For Each COLUMN_NAME As String In New String() _
                            {"INV_BALANCE", "INV_BALANCE_NEW", "INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF"}
                            .Item(COLUMN_NAME & "_CURR") = .Item(COLUMN_NAME)
                        Next
                        dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
                    End With

                    TOTAL_AR += INV_BALANCE

                    With rowARTOPEN1
                        .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT
                        .Item("INV_WRITE_OFF") = Val(.Item("INV_WRITE_OFF") & "") + INV_WRITE_OFF
                        .Item("INV_BALANCE") = INV_BALANCE - INV_PMT - INV_WRITE_OFF
                        .Item("INV_LAST_PMT") = rowARTPYMT1.Item("PYMT_BATCH_DATE")
                        .Item("INV_LAST_PMT_REF") = rowARTPYMT1.Item("PYMT_BATCH_NO")
                        .Item("INV_LAST_PMT_REF_DT") = rowARTPYMT1.Item("PYMT_BATCH_DATE")
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP

                        For Each COLUMN_NAME As String In New String() _
                            {"INV_BALANCE", "INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF"}
                            .Item(COLUMN_NAME & "_CURR") = .Item(COLUMN_NAME)
                        Next
                    End With

                    'If INV_BALANCE <> INV_PMT Then
                    '    Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                    '    With rowARTPYMT5
                    '        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                    '        .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO                           
                    '        .Item("PYMT_BATCH_DLNO") = 1
                    '        .Item("REASON_CODE") = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_CC_DIFF")
                    '        .Item("GL_DIST_AMT") = INV_BALANCE - INV_PMT
                    '        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    '        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    '        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    '        .Item("GL_DIST_AMT_CURR") = INV_BALANCE - INV_PMT      
                    '        dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
                    '    End With
                    'End If
                End If
            End If
        Next

        If CC_FEE <> 0 Then
            Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
            With rowARTPYMT2
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                PYMT_BATCH_LNO += 1
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                .Item("CUST_CODE") = DBNull.Value
                .Item("CUST_NAME") = DBNull.Value
                .Item("CUST_PYMT_REF_NO") = ""
                .Item("CUST_PYMT_REF_DATE") = DATETIME_STAMP.Date
                .Item("CUST_PYMT_AMT") = -1 * CC_FEE
                .Item("PYMT_STATUS") = "2"
                .Item("CUST_PYMT_AMT_CURR") = -1 * CC_FEE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("PYMT_NOTE") = "Credit Card Fee"
            End With
            dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

            Dim rowARTPYMT4 As DataRow = dst.Tables("ARTPYMT4").NewRow
            With rowARTPYMT4
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                .Item("PYMT_BATCH_GLNO") = 1
                .Item("GL_DIST_AMT") = CC_FEE
                .Item("GL_DIST_REF") = DBNull.Value
                .Item("GL_DIST_AUTO") = "1"
                .Item("ACCT_CODE") = ROWs("ARTPARM1").Item("AR_PARM_ACCT_CODE_CC_FEE")
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                If ASCMAIN1.CLIENT = "AHA" Then
                    .Item("SEG2_CODE") = "ECO"
                End If
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("GL_DIST_AMT_CURR") = CC_FEE
            End With
            dst.Tables("ARTPYMT4").Rows.Add(rowARTPYMT4)
        End If

        BeginTrans()
        Update_Record_TDA("ARTOPEN1")
        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTPYMT4")
        Update_Record_TDA("ARTPYMT5")
        Update_Record_TDA("ARTCCTR2")

        ASCMAIN1.sql = "Update ARTCCTR1 SET STATUS = '1' where STATUS = '0'" & vbCrLf _
            & " and (Select COUNT (*) FROM ARTCCTR2 WHERE CCTB_NO = ARTCCTR1.CCTB_NO AND PYMT_BATCH_NO IS NULL) = 0"
        ASCDATA1.ExecuteSQL()

        For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Select("")
            CUST_CODE = rowARTPYMT2.Item("CUST_CODE") & ""
            If CUST_CODE <> "" Then
                Dim i As Integer = Val(rowARTPYMT2.Item("PYMT_BATCH_LNO") & "")
                ASCDATA1.ExecuteSP("ARPPYMTP", "VN" _
                       , New Object() {PYMT_BATCH_NO, i} _
                       , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})

                ASCDATA1.ExecuteSP("ARPCUST6_PYMT", "VN" _
                                   , New Object() {PYMT_BATCH_NO, i} _
                                   , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})
            End If
        Next

        CommitTrans("Update Complete")
    End Sub

    Sub Refresh_Open_AR()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing AR Balances")

        ASCDATA1.ExecuteSQL("Update ARTCCTR2 Set INV_BALANCE = " _
                            & "(Select INV_BALANCE from ARTOPEN1 " _
                            & " where CUST_CODE = ARTCCTR2.CUST_CODE " _
                            & "   and INV_TYPE = ARTCCTR2.INV_TYPE " _
                            & "   and INV_NUM = ARTCCTR2.INV_NUM) where PYMT_BATCH_NO is Null")

        ASCMAIN1.sql = "Select * from ARTOPEN1 where (CUST_CODE,INV_TYPE,INV_NUM)" _
            & " in (Select CUST_CODE, INV_TYPE, INV_NUM from ARTCCTR2 where PYMT_BATCH_NO is Null)"
        Fill_Records("ARTOPEN1", "", True, ASCMAIN1.sql)

        Fill_Records("ARTCCTRX")

        For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("ISNULL(INV_NUM,'?') <> '?'")
            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                                         ({rowARTCCTR2.Item("CUST_CODE"), _
                                           rowARTCCTR2.Item("INV_TYPE"), _
                                           rowARTCCTR2.Item("INV_NUM")})
            If rowARTOPEN1 IsNot Nothing Then
                rowARTCCTR2.Item("INV_BALANCE") = rowARTOPEN1.Item("INV_BALANCE")
            Else
                rowARTCCTR2.Item("CUST_CODE") = DBNull.Value
                rowARTCCTR2.Item("INV_TYPE") = DBNull.Value
                rowARTCCTR2.Item("INV_NUM") = DBNull.Value
                rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = DBNull.Value
                rowARTCCTR2.Item("INV_BALANCE") = DBNull.Value
            End If
        Next

        For Each rowARTCCTR2 As DataRow In dst.Tables("ARTCCTR2").Select("ISNULL(INV_NUM,'?') = '?'")
            rowARTCCTR2.Item("MISSING_ORDER_NUMBER") = "1"
        Next

        grdARTCCTR2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Report()

    End Sub

    Sub Match_AR_Item(TRANS_AMT As Decimal, rowARTCCTR2 As DataRow)

        Dim sqlT As String = " and INV_TYPE = '" & IIf(TRANS_AMT >= 0, "I", "C") & "'"
        ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_CUST_PO = :PARM1" & sqlT

        ' to find the correct invoice number, if there was a reversal
        ASCMAIN1.sql &= " and INV_NO_REV is Null and INV_NO_REV_BY is Null order by ABS(INV_TOTAL_AMOUNT - " & CStr(TRANS_AMT) & ")"

        Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", rowARTCCTR2.Item("ORDER_NUMBER") & "")
        If rowSOTINVH1 IsNot Nothing Then Finalize_Match(rowSOTINVH1, rowARTCCTR2)

        If rowSOTINVH1 Is Nothing Then
            ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO_WEB = :PARM1" & sqlT
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", rowARTCCTR2.Item("ORDER_NUMBER") & "")
            If rowSOTINVH1 IsNot Nothing Then Finalize_Match(rowSOTINVH1, rowARTCCTR2)
        End If

        ' Use Transaction Number in ARTCCPA1 as a hit for TRADE charges
        If rowSOTINVH1 Is Nothing Then
            ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO = (Select ORDR_NO from ARTCCPA1 where TRANS_NUM = :PARM1) " & sqlT
            Dim ORDER_NUMBER As String = rowARTCCTR2.Item("ORDER_NUMBER") & ""
            If ORDER_NUMBER.Length < 10 Then ORDER_NUMBER = ORDER_NUMBER.PadLeft(10, "0")
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ORDER_NUMBER)
            If rowSOTINVH1 IsNot Nothing Then Finalize_Match(rowSOTINVH1, rowARTCCTR2)
        End If

        ' Returns use the Authorize.net Transaction Number
        If rowSOTINVH1 Is Nothing Then
            ASCMAIN1.sql = "Select * from SOTINVH1 where CC_CRED_TRANS_ID = :PARM1 " & sqlT
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", rowARTCCTR2.Item("ORDER_NUMBER") & "")
            If rowSOTINVH1 IsNot Nothing Then Finalize_Match(rowSOTINVH1, rowARTCCTR2)
        End If

        If rowSOTINVH1 Is Nothing And TRANS_AMT < 0 Then
            sqlT = " and INV_TYPE = 'I'"
            ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_CUST_PO = :PARM1" & sqlT
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", rowARTCCTR2.Item("ORDER_NUMBER") & "")
            If rowSOTINVH1 Is Nothing Then
                ASCMAIN1.sql = "Select * from SOTINVH1 where ORDR_NO_WEB = :PARM1" & sqlT
                rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", rowARTCCTR2.Item("ORDER_NUMBER") & "")
            End If
            If rowSOTINVH1 IsNot Nothing Then Finalize_Match(rowSOTINVH1, rowARTCCTR2)
        End If

        ' Look to see if the sales order is not finailzed.
        ' Arnold wants to see all PayPal transactions that have a status of Completed 
        If isProcessingPayPal AndAlso rowSOTINVH1 Is Nothing AndAlso (rowARTCCTR2.Item("ORDER_NUMBER") & String.Empty).ToString.Trim.Length > 0 Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_CUST_PO = :PARM1"
            Dim ORDER_NUMBER As String = (rowARTCCTR2.Item("ORDER_NUMBER") & String.Empty).ToString.Trim
            Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ORDER_NUMBER)
            If rowSOTORDR1 IsNot Nothing Then
                rowARTCCTR2.Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
            Else
                ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO_WEB = :PARM1"
                rowSOTORDR1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", ORDER_NUMBER)
                If rowSOTORDR1 IsNot Nothing Then
                    rowARTCCTR2.Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
                End If
            End If

            ' Assume All PayPal tranasaciotns are for consumer
            If rowARTCCTR2.Item("CUST_CODE") & "" = "" Then
                rowARTCCTR2.Item("CUST_CODE") = "CONSUMER"
            End If

            rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = 0
            rowARTCCTR2.Item("INV_BALANCE") = 0
        ElseIf isProcessingPayPal AndAlso rowSOTINVH1 Is Nothing AndAlso (rowARTCCTR2.Item("ORDER_NUMBER") & String.Empty).ToString.Trim.Length = 0 Then
            ' Assume All PayPal tranasaciotns are for consumer
            If rowARTCCTR2.Item("CUST_CODE") & "" = "" Then
                rowARTCCTR2.Item("CUST_CODE") = "CONSUMER"
            End If

            rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = 0
            rowARTCCTR2.Item("INV_BALANCE") = 0
        End If
    End Sub

    Sub Finalize_Match(ByRef rowSOTINVH1 As DataRow, rowARTCCTR2 As DataRow)

        If rowSOTINVH1 IsNot Nothing Then
            rowARTCCTR2.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
            rowARTCCTR2.Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE")
            rowARTCCTR2.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
            rowARTCCTR2.Item("INV_FREIGHT") = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
            rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            Dim rowARTOPEN1 As DataRow = LookUp("ARTOPEN1", New String() {rowARTCCTR2.Item("CUST_CODE"), rowARTCCTR2.Item("INV_TYPE"), rowARTCCTR2.Item("INV_NUM")})
            If rowARTOPEN1 IsNot Nothing Then
                ' rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = rowARTOPEN1.Item("INV_TOTAL_AMOUNT")
                rowARTCCTR2.Item("INV_BALANCE") = rowARTOPEN1.Item("INV_BALANCE")
            Else
                'ASCMAIN1.sql = "Select * from ARTOPENX where CUST_CODE = :PARM1 and INV_TYPE = :PARM2 and INV_NUM = :PARM3"
                Dim rowARTOPENX As DataRow = LookUp("ARTOPENX", New String() {rowARTCCTR2.Item("CUST_CODE"), rowARTCCTR2.Item("INV_TYPE"), rowARTCCTR2.Item("INV_NUM")})
                rowARTOPENX = Nothing ' BEFORE ALLOWING THIS TO HAPPEN YOU NEED TO SUPPORT THE REVIVAL FROM THE DEAD IN UPDATE DEPOSIT - BECAUSE AS OF 01/24/13 IT DOESN'T REVIVE FROM THE DEAD
                If rowARTOPENX IsNot Nothing Then
                    rowARTCCTR2.Item("INV_BALANCE") = 0
                Else
                    rowARTCCTR2.Item("CUST_CODE") = DBNull.Value
                    rowARTCCTR2.Item("INV_TYPE") = DBNull.Value
                    rowARTCCTR2.Item("INV_NUM") = DBNull.Value
                    rowARTCCTR2.Item("INV_TOTAL_AMOUNT") = DBNull.Value
                    rowSOTINVH1 = Nothing
                End If
            End If
        End If
    End Sub

    Private Sub GetPayPalData()

        Try
            FILENAME = String.Empty

            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select a PayPal text file to Import"
                openFileDialog1.Filter = "txt files (*.txt)|*.txt|Comma Separated (*.csv)|*.csv"
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    FILENAME = openFileDialog1.FileName
                End If
            End Using

            If FILENAME.Length = 0 Then
                EMsg &= vbCr & "File Selection Aborted."
                Exit Sub
            End If

            Dim iline As Int16 = 0
            Dim inputLine As String = String.Empty

            Dim fileExtension As String = System.IO.Path.GetExtension(FILENAME).ToUpper
            Dim isCSVFile As Boolean = fileExtension = ".CSV"

            If isCSVFile Then
                Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(FILENAME)

                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    MyReader.SetDelimiters(",")

                    Dim currentRow As String()

                    While Not MyReader.EndOfData
                        currentRow = MyReader.ReadFields()

                        If currentRow.Length < 3 Then
                            Continue While
                        End If

                        'If currentRow(2) & String.Empty <> "Time Zone" AndAlso iline = 0 Then
                        '    Continue While
                        'End If
                        If (currentRow(2) & String.Empty <> "Time Zone" And currentRow(2) & String.Empty <> "TimeZone") AndAlso iline = 0 Then
                            Continue While
                        End If

                        iline += 1

                        If iline = 1 Then
                            dt = New DataTable("PayPal")
                            For Each fieldName As String In currentRow
                                dt.Columns.Add(fieldName.Trim.Replace(" ", "_").Replace("/", "_"))
                            Next
                            Continue While
                        End If

                        dt.Rows.Add(currentRow)

                    End While
                    MyReader.Close()
                    MyReader.Dispose()
                End Using

                'Using sr As New System.IO.StreamReader(FILENAME)
                '    Do While sr.Peek() > -1
                '        inputLine = sr.ReadLine

                '        If Not inputLine.Contains("Time Zone") AndAlso iline = 0 Then
                '            Continue Do
                '        End If

                '        iline += 1
                '        If iline = 1 Then
                '            dt = New DataTable("PayPal")
                '            For Each fieldName As String In inputLine.Split(",")
                '                dt.Columns.Add(fieldName.Trim.Replace(" ", "_").Replace("/", "_"))
                '            Next
                '            Continue Do
                '        End If

                '        inputLine = inputLine.Replace(Chr(34) & ",," & Chr(34), Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34))
                '        If inputLine.Contains(Chr(34) & "," & Chr(34)) Then
                '            inputLine = inputLine.Replace(Chr(34) & "," & Chr(34), ControlChars.Tab)
                '        Else
                '            inputLine = inputLine.Replace(",", ControlChars.Tab)
                '        End If
                '        inputLine = inputLine.Replace(Chr(34), "")

                '        dt.Rows.Add(inputLine.Split(ControlChars.Tab))

                '    Loop
                '    sr.Close()
                '    sr.dispose
                'End Using
            Else
                Using sr As New System.IO.StreamReader(FILENAME)
                    Do While sr.Peek() > -1
                        inputLine = sr.ReadLine
                        inputLine = inputLine.Replace(Chr(34), "")

                        iline += 1
                        If iline = 1 Then
                            dt = New DataTable("PayPal")
                            For Each fieldName As String In inputLine.Split(ControlChars.Tab)
                                dt.Columns.Add(fieldName.Trim.Replace(" ", "_").Replace("/", "_"))
                            Next
                            Continue Do
                        End If

                        dt.Rows.Add(inputLine.Split(ControlChars.Tab))

                    Loop
                    sr.Close()
                    sr.Dispose()
                End Using
            End If


            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                EMsg &= vbCr & "No data found to process."
            End If

        Catch ex As Exception
            EMsg &= vbCr & ex.Message
        End Try

    End Sub

    Private Sub ProcessPayPalData()

        ' 10/29/2015
        ' Arnold wants to see all transactions that have a status of Completed.

        Dim r As Integer = 0
        Dim r_start As Integer = 0
        Dim CCTB_LNO As Integer = 0
        Dim rowARTCCPA1 As DataRow = Nothing

        rowARTCCTR1.Item("CCTB_DATE") = DateTime.Now.ToShortDateString
        rowARTCCTR1.Item("CCTB_MEMO") = "Paypal: " & DateTime.Now.ToString("yyyyMMdd")
        rowARTCCTR1.Item("TOTAL_UNMATCHED") = 0

        For Each row As DataRow In dt.Rows

            Dim gross As Decimal = Val((row.Item("Gross") & String.Empty).ToString.Replace(",", ""))
            Dim Net As Decimal = Val((row.Item("Net") & String.Empty).ToString.Replace(",", ""))
            Dim serviceFee As Decimal = Val((row.Item("Fee") & String.Empty).ToString.Replace(",", ""))

            If gross = 0 AndAlso Net = 0 Then
                Continue For
            End If

            Dim rowARTCCTR2 As DataRow = dst.Tables("ARTCCTR2").NewRow
            rowARTCCTR2.Item("CCTB_NO") = HFs("CCTB_NO")
            CCTB_LNO += 1
            rowARTCCTR2.Item("CCTB_LNO") = CCTB_LNO

            Dim ORDR_NO_WEB As String = row.Item("Invoice_Number") & String.Empty
            ORDR_NO_WEB = ORDR_NO_WEB.Trim
            ASCMAIN1.Progress("-", ORDR_NO_WEB)

            Dim TRANS_NAME As String = "Sale"

            If gross < 0 Then
                TRANS_NAME = "Return"
            End If

            rowARTCCTR2.Item("CARD_NAME") = String.Empty
            rowARTCCTR2.Item("TRANS_NAME") = TRANS_NAME
            If IsDate((row.Item("Date") & " " & row.Item("Time")).ToString.Trim) Then
                rowARTCCTR2.Item("TIME") = CDate((row.Item("Date") & " " & row.Item("Time")).ToString.Trim)
            End If
            rowARTCCTR2.Item("ENTRY_MODE") = row.Item("Type") & String.Empty
            rowARTCCTR2.Item("AUTH_CODE") = String.Empty
            rowARTCCTR2.Item("TRANS_AMT") = gross

            rowARTCCTR2.Item("SERVICE_FEE") = serviceFee
            rowARTCCTR2.Item("TRANS_AMT_NET") = Net

            rowARTCCTR2.Item("TRANSACTION") = row.Item("Status") & String.Empty
            rowARTCCTR2.Item("CARD_NUM") = DBNull.Value
            rowARTCCTR2.Item("VOID") = DBNull.Value
            rowARTCCTR2.Item("REFERENCE_NUMBER") = row.Item("Transaction_ID") & String.Empty
            rowARTCCTR2.Item("EXPIRATION_DATE") = DBNull.Value
            rowARTCCTR2.Item("AUTH_SOURCE") = row.Item("Type") & String.Empty
            rowARTCCTR2.Item("CARD_METHOD") = DBNull.Value
            rowARTCCTR2.Item("POS_CAPABILITY") = DBNull.Value
            rowARTCCTR2.Item("AVS") = row.Item("Address_Status") & String.Empty
            rowARTCCTR2.Item("CVV2") = DBNull.Value
            rowARTCCTR2.Item("TIP_AMT") = 0
            rowARTCCTR2.Item("ORDER_NUMBER") = row.Item("Invoice_Number") & String.Empty
            rowARTCCTR2.Item("RETURN_COUNT") = DBNull.Value

            For I As Integer = 0 To 18
                'rowARTCCTR2.Item(I + 2) = row.Item(I)

                With dst.Tables("ARTCCTR2").Columns(I + 2)
                    If .DataType = GetType(System.String) Then
                        If Len(rowARTCCTR2.Item(I + 2) & "") > .MaxLength Then
                            rowARTCCTR2.Item(I + 2) = Mid((rowARTCCTR2.Item(I + 2) & ""), 1, .MaxLength)
                        End If
                    End If
                End With
            Next

            Dim skip_record As Boolean = False
            'If rowARTCCTR2.Item("CARD_NAME") & "" = "" Then
            '    If rowARTCCTR2.Item("TRANS_NAME") & "" = "" And _
            '        rowARTCCTR2.Item("CARD_NUM") & "" = "" And _
            '        rowARTCCTR2.Item("AUTH_CODE") & "" = "" And _
            '        rowARTCCTR2.Item("EXPIRATION_DATE") & "" = "" And _
            '        rowARTCCTR2.Item("TRANS_AMT") & "" = "" Then
            '        skip_record = True
            '    End If
            'End If

            ' Requested by Arnold 12/15/2015
            If (rowARTCCTR2.Item("ENTRY_MODE") & String.Empty).ToString.ToUpper.StartsWith("Transfer to Bank".ToUpper) Then
                skip_record = True
            End If


            If Not skip_record Then
                dst.Tables("ARTCCTR2").Rows.Add(rowARTCCTR2)

                ' AR Item Matching Process
                Dim TRANS_AMT As Decimal = Val(rowARTCCTR2.Item("TRANS_AMT") & "")
                Match_AR_Item(TRANS_AMT, rowARTCCTR2)
            End If
        Next

        For Each row As DataRow In dst.Tables("ARTCCTR2").Select("TRANSACTION <> 'Completed'")
            ' After reconciling October & November I learned that transactions with a “Denied” or “Pending Status” 
            ' with type “Funds Released” or “Funds Held” are also part of the PayPal balance. 
            ' These transactions result from consumer disputes & they don’t get reported as “Complete” after 
            ' resolution (See attached examples). Accordingly, they should also be included in the import.
            If ((row.Item("ENTRY_MODE") & String.Empty).ToString.ToUpper.StartsWith("Funds Released".ToUpper) _
                OrElse (row.Item("ENTRY_MODE") & String.Empty).ToString.ToUpper.StartsWith("Transfer to Bank".ToUpper)) _
                AndAlso ((row.Item("TRANSACTION") & String.Empty).ToString.ToUpper.StartsWith("Denied".ToUpper) _
                OrElse (row.Item("TRANSACTION") & String.Empty).ToString.ToUpper.StartsWith("Pending".ToUpper)) Then
                ' Keep the record
            Else
                rowARTCCTR1.Item("TOTAL_UNMATCHED") += Val(row.Item("TRANS_AMT") & String.Empty)
                row.Delete()
            End If
        Next

    End Sub

    Private Sub grdARTOPENX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTOPENX.DoubleClickRow

        Dim ORDR_NO As String = e.Row.Cells("INV_CUST_PO").Value

        ASCMAIN1.sql = "Select .* from ARTCCTR1, ARTCCTR2 " _
            & " where ARTCCTR1.CCTB_NO = ARTCCTR2.CCTB_NO" _
            & " AND ARTCCTR2.ORDER_NUMBER = '" & ORDR_NO & "' and ARTCCTR2.TRANS_AMT > 0" _
            & " AND ARTCCTR1.STATUS <> 'D'"
        Dim rowARTCCTR2x As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
        If Not (rowARTCCTR2x Is Nothing AndAlso dst.Tables("ARTCCTR2").Select("ORDER_NUMBER = '" & ORDR_NO & "' And TRANS_AMT > 0").Length = 0) Then
            MessageBox.Show("The selected Open AR Entry is already matched to an existing entry.")
        Else
            Dim CCTB_NO As String = grdARTCCTR2.ActiveRow.Cells("CCTB_NO").Value
            Dim CCTB_LNO As Int32 = grdARTCCTR2.ActiveRow.Cells("CCTB_LNO").Value

            Dim rowARTCCTR2 As DataRow = dst.Tables("ARTCCTR2").Rows.Find(New Object() {CCTB_NO, CCTB_LNO})
            If rowARTCCTR2 Is Nothing Then
                MessageBox.Show("Could not locate the entry to Update.")
                Exit Sub
            End If

            rowARTCCTR2.Item("ORDER_NUMBER") = ORDR_NO
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = :PARM1 AND INV_NO = :PARM2", "VV", New Object() {e.Row.Cells("INV_TYPE").Value, e.Row.Cells("INV_NUM").Value})
            Finalize_Match(rowSOTINVH1, rowARTCCTR2)
        End If

    End Sub

    Private Sub grdARTOPENX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPENX.InitializeRow
        Dim TRANS_AMT As Decimal = Val(grdARTCCTR2.ActiveRow.Cells("TRANS_AMT").Value & String.Empty)

        If (e.Row.Cells("INV_BALANCE").Value & String.Empty) <> TRANS_AMT Then
            e.Row.Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub chkUnappliedOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkUnappliedOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Load_ARTCCTRX(True)
    End Sub
End Class