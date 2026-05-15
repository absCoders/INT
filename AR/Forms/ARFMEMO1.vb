Public Class ARFMEMO1

    ' PICK FROM A LIST OF INVOICES (ESP USEFUL FOR WEB CREDITS)
    ' IF WEB CREDIT, SEND CR MEMO 
    Dim rowSOTINVH1 As DataRow
    Dim rowSOTINVHC As DataRow

    Dim auto_CR As Boolean

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")
        If MENU_ITEM_OBJECT = "ARFMEMOI" Then
            InquiryMode = True
        End If

        With dst
            Create_TDA(.Tables.Add, "SOTINVH1", "*")

            ASCMAIN1.sql = "Select INV_TYPE, INV_NUM, INV_DATE, INV_BALANCE, '0' SEL" _
            & " from ARTOPEN1 where CUST_CODE = :PARM1 and INV_BALANCE <> 0"
            Call Create_TDA(.Tables.Add, "ARTOPENA", "**", 0, False, "V", 0)
            .Tables("ARTOPENA").Columns.Add("INV_BALANCE_APPLIED", GetType(System.Decimal))

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")


            ASCMAIN1.sql = "Select INV_TYPE, INV_NO, INV_DATE, CUST_CODE, MISC_CHG_CODE, INIT_OPER, INIT_DATE" _
                & ", ORDR_CUST_PO, SALES_DIVISION_CODE, INV_COMMENT" _
                & ", ORDR_TYPE_CODE" _
                & ", INV_SALES, INV_FREIGHT, INV_MISC_CHG, INV_STAX, INV_TOTAL_AMOUNT" _
                & " from SOTINVH1 where ORDR_TYPE_CODE = 'TOP' and ORDR_YYYYPP_UPDATED = :PARM1"
            Call Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "V", 2)
            '  .Tables("SOTINVHX").Columns.Add("INV_BALANCE_APPLIED", GetType(System.Decimal))



            ASCMAIN1.sql = "Select INV_TYPE, INV_NO, INV_DATE, CUST_CODE, MISC_CHG_CODE, INIT_OPER, INIT_DATE" _
                & ", ORDR_CUST_PO, SALES_DIVISION_CODE, INV_COMMENT" _
                & ", ORDR_TYPE_CODE, CCPA_NO, CC_SALE_TRANS_ID" _
                & ", INV_SALES, INV_FREIGHT, INV_MISC_CHG, INV_STAX, INV_TOTAL_AMOUNT" _
                & " from SOTINVH1 where INV_TYPE = 'I' and CUST_CODE = :PARM1 and ORDR_YYYYPP_UPDATED = :PARM1"
            Call Create_TDA(.Tables.Add, "SOTINVHC", "**", 0, False, "VV", 2)

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")
            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
        End With


        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)


        grdARTOPENA.DataSource = dst.Tables("ARTOPENA")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdSOTINVHC.DataSource = dst.Tables("SOTINVHC")

        Create_Summary(grdARTOPENA, "INV_TYPE", "Count")
        Create_Summary(grdARTOPENA, "INV_BALANCE")
        Create_Summary(grdARTOPENA, "INV_BALANCE_APPLIED")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "INV_FREIGHT")
        '   Create_Summary(grdSOTINVHX, "INV_STAX")
        Create_Summary(grdSOTINVHX, "INV_MISC_CHG")
        Create_Summary(grdSOTINVHX, "INV_TOTAL_AMOUNT")

        Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)

        Show_Filter(grdSOTINVHC, True)

        Bind_Controls(grpSOTINVHC, "SOTINVHC")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("CUST_CODE")

                If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If

                'If Absx1.optFor("INV_TYPE").Value & "" = "" Then
                '    EMsg &= vbCr & "Memo Type Required"
                'End If

                ' MULTITASKING

            Case "View"
                'Validate_Code("INV_NO")
                If Absx1.txtFor("INV_NO").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Document No"
                Else
                    ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & Absx1.txtFor("INV_NO").Text & "'"
                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
                    If rowSOTINVH1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Document No"
                    Else
                        If rowSOTINVH1.Item("ORDR_TYPE_CODE") <> "TOP" Then
                            EMsg &= vbCr & "Invalid Type of Document to view with this screen"
                        Else
                            Absx1.optFor("INV_TYPE").Value = rowSOTINVH1.Item("INV_TYPE")
                        End If
                    End If
                End If

            Case "Reverse"

                Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")
                Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
                Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")

                If MsgBox("Do you want to set up an entry to Reverse " & IIf(INV_TYPE = "C", "CR", "DR") & " Memo " & INV_NO, vbYesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If Val(Absx1.numFor("INV_STAX").Value & "") <> 0 Then
                    If Absx1.txtFor("STAX_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Tax Code"
                    Else
                        Validate_Code("STAX_CODE")
                    End If
                End If

                If Val(Absx1.numFor("INV_MISC_CHG").Value & "") <> 0 Then
                    If Absx1.txtFor("MISC_CHG_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Misc Chg Code"
                    Else
                        Validate_Code("MISC_CHG_CODE")
                    End If

                End If

                ' #ISSUE-7388 Call w/DH/AK to allow specific customers to optionally have a Colection Code specified
                Dim CUST_CODEs As New List(Of String) ' customers who should not require a collection code to be specified
                CUST_CODEs.Add("IPUSA")
                CUST_CODEs.Add("IPSA")
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

                If CUST_CODEs.Contains(CUST_CODE) Then
                    If Absx1.txtFor("COLLECTION_CODE").Text <> "" Then
                        Validate_Code("COLLECTION_CODE")
                    End If
                Else
                    If Absx1.txtFor("COLLECTION_CODE").Text = "" Then
                        ' ISSUE-6875 9/18/25: Collection is mandatory, removed warning
                        EMsg &= vbCr & "You Must Specify a Collection Code"
                    Else
                        Validate_Code("COLLECTION_CODE")
                    End If

                End If



                If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") = 0 Then
                    EMsg &= vbCr & "Total Amount is Zero"
                Else
                    If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") > 0 And Absx1.optFor("INV_TYPE").Value = "C" Then
                        EMsg &= vbCr & "Credit Amount may not be Positive"
                    End If
                    If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") < 0 And Absx1.optFor("INV_TYPE").Value = "D" Then
                        EMsg &= vbCr & "Debit Amount may not be Negative"
                    End If
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                        ' OK FOR INT
                    Else
                        Validate_Code("SALES_DIVISION_CODE")
                    End If

                Else
                    Validate_Code("SALES_DIVISION_CODE")
                End If

                Dim rowSOTMISC1 As DataRow = LookUp("SOTMISC1", Absx1.txtFor("MISC_CHG_CODE").Text)
                If Absx1.txtFor("REASON_CODE").Text = "" Then
                    If rowSOTMISC1 IsNot Nothing AndAlso rowSOTMISC1.Item("REASON_CODE") & "" <> "" Then
                        If ASCMAIN1.CLIENT = "INT" Then
                            ' INT DOES NOT CORRELATE REASON CODES TO MISC CHARGES
                            EMsg &= vbCr & "Reason Code is Mandatory"
                        Else
                            ' OK - WE WILL USE REASON CODE ASSOCIATED WITH MISC CHG CODE
                        End If
                    Else
                        EMsg &= vbCr & "Reason Code is Mandatory (when not using a Corrleated Misc Chg Code)"
                    End If
                Else
                    Validate_Code("REASON_CODE")
                    If rowSOTMISC1 IsNot Nothing Then
                        If Absx1.txtFor("REASON_CODE").Text <> rowSOTMISC1.Item("REASON_CODE") & "" And rowSOTMISC1.Item("REASON_CODE") & "" <> "" Then
                            EMsg &= vbCr & "Reason Code does not match Reason associated with Misc Chg Code"
                        End If
                    End If
                End If

                Dim DT As Date = Absx1.dteFor("INV_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                If EMsg = "" Then
                    If chkIssueCCCredit.Visible And Not chkIssueCCCredit.Checked Then
                        If MsgBox("The Customer's Terms are Credit Card" _
                                  & vbCrLf & "Yet this Credit is not indicated to be Credited via Credit Card" _
                                  & vbCrLf & vbCrLf & "Continue with Update?",
                                  MsgBoxStyle.YesNo, "Option to Credit Customer's Credit Card") <> MsgBoxResult.Yes Then
                            Exit Sub
                        End If
                    End If
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Reverse"

                Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")
                Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
                Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")

                Mode_Settings(False)

                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                If INV_TYPE = "C" Then
                    Absx1.optFor("INV_TYPE").Text = "I"
                Else
                    Absx1.optFor("INV_TYPE").Text = "C"
                End If
                Click_Command("New")


                Dim rowSOTINVH1_to_reverse As DataRow = LookUp("SOTINVH1", New String() {INV_TYPE, INV_NO})

                For Each C As String In New String() _
                    {"ORDR_CUST_PO", "SALES_DIVISION_CODE", "COLLECTION_CODE", "MISC_CHG_CODE", "REASON_CODE"}
                    Absx1.txtFor(C).Text = rowSOTINVH1_to_reverse.Item(C) & ""
                    '  rowSOTINVH1.Item(C) = rowSOTINVH1_to_reverse.Item(C)
                Next

                Absx1.txtFor("INV_COMMENT").Text = "Reversal of " & IIf(INV_TYPE = "C", "CR", "DR") & " Memo " & INV_NO

                Absx1.numFor("INV_MISC_CHG").Value = -1 * Val(rowSOTINVH1_to_reverse.Item("INV_MISC_CHG"))
                'Absx1.txtFor("INV_TOTAL_AMOUNT").Text = -1 * Val(rowSOTINVH1_to_reverse.Item("INV_TOTAL_AMOUNT"))
                'rowSOTINVH1.Item("INV_MISC_CHG") = -1 * Val(rowSOTINVH1_to_reverse.Item("INV_MISC_CHG"))
                'rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = -1 * Val(rowSOTINVH1_to_reverse.Item("INV_TOTAL_AMOUNT"))

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print Credit Memo"
                Print_Credit_Memo()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    .Items("Reverse").Visible = ScreenMode And (EntryMode = "V") And Not InquiryMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    If InquiryMode Then
                        With UltraExplorerBar1.Groups("Screen Control")
                            .Items("New").Visible = False
                            .Items("Update").Visible = False
                        End With
                    End If
                End With

                .Groups("Show If Entered in").Visible = Not ScreenMode

                '(InquiryMode OrElse EntryMode = "V" OrElse ScreenMode)
                .Groups("Screen Control").Items("Print Credit Memo").Visible = (EntryMode = "V" And ScreenMode) _
                        AndAlso (rowSOTINVH1 IsNot Nothing)
                If ScreenMode Then
                    With .Groups("Screen Control").Items("Print Credit Memo")
                        If (Absx1.optFor("INV_TYPE").Value = "C") Then
                            .Text = "Print CR Memo"
                        Else
                            .Text = "Print DR Memo"
                        End If
                    End With

                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpMemoDetails.Visible = tf
        ' grdSOTINVHX.Visible = Not tf

        grdARTOPENA.Visible = False
        tab0.Visible = Not tf

        lblNote.Visible = ScreenMode And (Absx1.optFor("INV_TYPE").Value = "C")

        Set_Read_Only(grpMemoDetails, (EntryMode = "V"))
        'Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)
        If ScreenMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), auto_CR)
            grpSOTINVHC.Visible = auto_CR
        End If
        Set_Read_Only(grpSOTINVHC, True)

        lblDRCR.Text = Absx1.optFor("INV_TYPE").Text
        If Absx1.optFor("INV_TYPE").Value = "C" Then
            lblDRCR.Appearance.ForeColor = Drawing.Color.Red
        Else
            lblDRCR.Appearance.ForeColor = Drawing.Color.Green
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTINVH1", "ARTOPENA", "ARTPYMT1", "ARTPYMT2", "ARTPYMT3", "SOTINVHC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Setup_tab0()

        Absx1.dteFor("INV_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.optFor("INV_TYPE").Value = "C"
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowSOTINVH1 = dst.Tables("SOTINVH1").NewRow
            rowSOTINVH1.Item("INV_TYPE") = HFs("INV_TYPE")
            rowSOTINVH1.Item("INV_NO") = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            rowSOTINVH1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowSOTINVH1.Item("CUST_BILL_TO_CUST") = HFs("CUST_CODE")
            rowSOTINVH1.Item("CUST_STORE_NO") = "000000"
            rowSOTINVH1.Item("INV_DATE") = Now.Date ' HFs("INV_DATE")
            rowSOTINVH1.Item("ORDR_TYPE_CODE") = "TOP"
            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            rowSOTINVH1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
            rowSOTINVH1.Item("INV_SALES") = 0
            rowSOTINVH1.Item("INV_COGS") = 0
            rowSOTINVH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTINVH1.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTINVH1.Item("INV_PRINTED") = "1"
            rowSOTINVH1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            rowSOTINVH1.Item("CURR_EXCH_RATE") = 1
            rowSOTINVH1.Item("REGISTER_IND") = "0"

            If auto_CR Then
                rowSOTINVH1.Item("ORDR_CUST_PO") = rowSOTINVHC.Item("ORDR_CUST_PO")
                rowSOTINVH1.Item("SALES_DIVISION_CODE") = rowSOTINVHC.Item("SALES_DIVISION_CODE")
            End If

            ' Place the original Invoice Number and CC Trans ID from the Invoice on the credit
            If rowSOTINVH1.Item("INV_TYPE") = "C" Then
                If rowSOTINVHC IsNot Nothing AndAlso rowSOTINVHC.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty Then
                    rowSOTINVH1.Item("CC_SALE_TRANS_ID") = rowSOTINVHC.Item("CC_SALE_TRANS_ID") & String.Empty
                    rowSOTINVH1.Item("INV_NO_CR") = rowSOTINVHC.Item("INV_NO") & String.Empty
                ElseIf MyBase.Absx1.txtFor("CCPA_NO").Text.Trim.Length > 0 Then
                    Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & MyBase.Absx1.txtFor("CCPA_NO").Text & "'")
                    If rowARTCCPA1 IsNot Nothing Then
                        rowSOTINVH1.Item("CC_SALE_TRANS_ID") = rowARTCCPA1.Item("TRANS_ID") & String.Empty
                        rowSOTINVH1.Item("INV_NO_CR") = txtInvNo.Text
                    End If
                End If
            End If

            'remaining SOTINVH1 fields
            'CUST_SHIP_TO_NO
            'ORDR_NO
            'WHSE_CODE
            'REGISTER_XNO
            'SHIPMENT_NO
            'PICK_NO
            'CUST_SHIP_TO_STATE
            'REGISTER_IND
            'APPLY_TO_INV_TYPE
            'APPLY_TO_INV_NO
            'INV_REVERSED
            'INV_REVERSED_INV_NO
            'INV_NO_RESHIP
            'INV_RESHIP

            'defaults from customer
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", New String() {HFs("CUST_CODE")}, True)
            rowSOTINVH1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            rowSOTINVH1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0") ' rowARTCUST1.Item("TERM_CODE")

            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1.ItemArray)

            Fill_Records("ARTOPENA", HFs("CUST_CODE"))

            chkIssueCCCredit.Visible = False
            chkIssueCCCredit.Checked = False

            Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", rowSOTINVH1.Item("TERM_CODE") & "")
            If rowTATTERM1 IsNot Nothing Then
                If rowTATTERM1.Item("TERM_TYPE") & "" = "D" Then
                    chkIssueCCCredit.Visible = True
                    ' chkIssueCCCredit.Checked = True
                End If
            End If
        Else
            ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & Absx1.txtFor("INV_NO").Text & "'"
            rowSOTINVH1 = Fill_Record("SOTINVH1", New String() {Absx1.optFor("INV_TYPE").Value, Absx1.txtFor("INV_NO").Text})
            dst.AcceptChanges()

            chkIssueCCCredit.Visible = False
        End If





        ' NEED TO FILL ARTOPENA DIFFERENTLY IF THIS FORM IS TO BE USED IN INQUIRY MODE

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        rowSOTINVH1 = dst.Tables("SOTINVH1").Rows(0)
        If rowSOTINVH1.Item("REASON_CODE") & "" = "" AndAlso rowSOTINVH1.Item("MISC_CHG_CODE") & String.Empty <> String.Empty Then
            Dim rowSOTMISC1 As DataRow = LookUp("SOTMISC1", rowSOTINVH1.Item("MISC_CHG_CODE") & String.Empty)
            If rowSOTMISC1 IsNot Nothing Then
                If rowSOTMISC1.Item("REASON_CODE") & "" <> "" Then
                    rowSOTINVH1.Item("REASON_CODE") = rowSOTMISC1.Item("REASON_CODE")
                End If
            End If
        End If

        ' See if we can find the original credit card transaction. the user may have just entered the Customer code then the Customer Reference
        If rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty = String.Empty Then

            MyBase.Absx1.txtFor("ORDR_CUST_PO").Text = MyBase.Absx1.txtFor("ORDR_CUST_PO").Text.Trim
            Dim ORDR_CUST_PO As String = MyBase.Absx1.txtFor("ORDR_CUST_PO").Text

            If ORDR_CUST_PO.Length > 0 Then
                Dim CUST_CODE As String = HFs("CUST_CODE")
                Dim INV_NO_CR As String = String.Empty
                Dim CC_SALE_TRANS_ID As String = String.Empty
                SOCMAIN1.GetCreditCardSaleTransaction(CUST_CODE, ORDR_CUST_PO, INV_NO_CR, CC_SALE_TRANS_ID)
                rowSOTINVH1.Item("CC_SALE_TRANS_ID") = CC_SALE_TRANS_ID
                rowSOTINVH1.Item("INV_NO_CR") = INV_NO_CR
            End If

        End If

        Update_Record_TDA("SOTINVH1")

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.Item("CUST_CODE") = HFs("CUST_CODE")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE")
        rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
        rowARTOPEN1.Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
        'rowARTOPEN1.Item("CUST_SHIP_TO_NO")
        rowARTOPEN1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE") ' "REG" 'should we allow post code on screen
        rowARTOPEN1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")
        rowARTOPEN1.Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
        rowARTOPEN1.Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE")
        rowARTOPEN1.Item("STAX_CODE") = rowSOTINVH1.Item("STAX_CODE")
        'rowARTOPEN1.Item("APPLY_TO_INV_NUM")
        'rowARTOPEN1.Item("APPLY_TO_INV_TYPE")
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
        rowARTOPEN1.Item("INV_SALES") = rowSOTINVH1.Item("INV_SALES")
        rowARTOPEN1.Item("INV_DISC") = 0
        rowARTOPEN1.Item("INV_FREIGHT") = rowSOTINVH1.Item("INV_FREIGHT")
        rowARTOPEN1.Item("INV_STAX") = rowSOTINVH1.Item("INV_STAX")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("REASON_CODE") = rowSOTINVH1.Item("REASON_CODE")
        rowARTOPEN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("INIT_DATE") = DATETIME_STAMP
        rowARTOPEN1.Item("INV_MISC_CHG") = rowSOTINVH1.Item("INV_MISC_CHG")
        rowARTOPEN1.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowARTOPEN1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        rowARTOPEN1.Item("CURR_EXCH_RATE") = 1
        rowARTOPEN1.Item("INV_SALES_CURR") = 0
        rowARTOPEN1.Item("INV_DISC_CURR") = 0
        rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT")
        rowARTOPEN1.Item("INV_STAX_CURR") = rowSOTINVH1.Item("INV_STAX")
        rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowSOTINVH1.Item("INV_MISC_CHG")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_BALANCE_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_NOTES") = rowSOTINVH1.Item("INV_COMMENT")
        rowARTOPEN1.Item("ORDR_TYPE_CODE") = rowSOTINVH1.Item("ORDR_TYPE_CODE")
        'rowARTOPEN1.Item("INV_REF") = rowSOTINVH1.Item("INV_REF")
        rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowARTOPEN1.Item("SALES_DIVISION_CODE") = rowSOTINVH1.Item("SALES_DIVISION_CODE")
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        ' APPLY TO OPEN AR ITEMS

        'If "THERE WAS SOMETHING APPLIED" Then

        ' ESTABLISH A SINGLE ROW IN ARTPYMT1
        ' ESTABLISH A SINGLE ROW IN ARTPYMT2

        '    Dim INV_PMT_TOTAL As Decimal = 0
        '    For Each rowARTOPENA As DataRow In dst.Tables("ARTOPENA").Select("SEL = '1'")
        '        rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() _
        '        {HFs("CUST_CODE"), _
        '        rowARTOPENA.Item("INV_TYPE"), _
        '        rowARTOPENA.Item("INV_NUM")}, False, False)

        '        Dim INV_BALANCE As Double = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
        '        Dim INV_PMT As Double = INV_BALANCE
        '        INV_PMT_TOTAL += INV_PMT
        '        ' WE WILL HAVE TO HANDLE THESE WHEN WE DO A BUYING GROUP WITH ANTIC DISC
        '        Dim INV_DISC_TAKEN As Double = 0
        '        Dim INV_WRITE_OFF As Double = 0

        '        Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2_BOX, _
        '        ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"), _
        '         PYMT_BATCH_DATE, INV_PMT, INV_DISC_TAKEN, INV_WRITE_OFF, PYMT_BATCH_ILNO)

        '    Next
        'End If

        Update_Record_TDA("ARTOPEN1")

        ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
                           New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")},
                           New String() {"INV_TYPE_IN", "INV_NO_IN"})

        CommitTrans("Update Complete")

        ' See if we need to issue credit card credit.
        If rowSOTINVH1.Item("INV_TYPE") = "C" AndAlso rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty Then
            ' This is done since paypal transaction IDs cannt be used for Authorize.net
            If MessageBox.Show("Do you want to refund the Credit Card?", "Refund", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                ASCMAIN1.Progress("Processing CC Credit", "")
                Dim errorMessage As String = String.Empty
                If Not SOCMAIN1.IssueCredit(rowSOTINVH1.Item("INV_NO"), errorMessage) Then
                    MessageBox.Show("Error Processing Credit Card Refund: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ASCMAIN1.Progress("", "")
            End If
        End If

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    ' Click_Command("New", e)
                    Load_SOTINVHC()
                End If
            Case "INV_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                '   Click_Command("New")
                Load_SOTINVHC()
            Case "INV_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        Call LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then
                            Load_SOTINVHC()
                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT", "INV_STAX", "INV_MISC_CHG"
                Absx1.numFor("INV_TOTAL_AMOUNT").Value = Val(Absx1.numFor("INV_FREIGHT").Value & "") +
                Val(Absx1.numFor("INV_STAX").Value & "") +
                Val(Absx1.numFor("INV_MISC_CHG").Value & "")
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTINVHX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            'Case "grdSOTRMAF2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

    Private Sub grdARTOPENA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPENA.AfterCellUpdate
        If e.Cell.Column.Key = "SEL" Then
            If e.Cell.Value = "1" Then
                e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = e.Cell.Row.Cells("INV_BALANCE").Value
            Else
                e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = 0
            End If
        End If
    End Sub

    Private Sub grdARTOPENA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTOPENA.BeforeRowUpdate
        'If e.Row.Cells("SEL").Value = "1" Then
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = e.Row.Cells("INV_BALANCE").Value
        'Else
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        'End If
    End Sub

    Sub Pay_Open_AR_Item(
    ByVal rowARTOPEN1 As DataRow,
    ByVal rowARTPYMT2 As DataRow,
    ByVal CURR_CODE As String,
    ByVal PYMT_BATCH_DATE As Date,
    ByVal INV_PMT As Double,
    ByVal INV_DISC_TAKEN As Double,
    ByVal INV_WRITE_OFF As Double,
    ByRef PYMT_BATCH_ILNO As Integer)

        With rowARTOPEN1
            Dim INV_BALANCE As Double = Val(.Item("INV_BALANCE") & "")

            .Item("INV_LAST_PMT") = PYMT_BATCH_DATE
            .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT
            .Item("INV_DISC_TAKEN") = Val(.Item("INV_DISC_TAKEN") & "") + INV_DISC_TAKEN
            .Item("INV_WRITE_OFF") = Val(.Item("INV_WRITE_OFF") & "") + INV_WRITE_OFF
            .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - (INV_PMT + INV_DISC_TAKEN + INV_WRITE_OFF)
            .Item("INV_LAST_PMT_REF") = rowARTPYMT2.Item("CUST_PYMT_REF_NO")
            .Item("INV_LAST_PMT_REF_DT") = rowARTPYMT2.Item("CUST_PYMT_REF_DATE")
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                .Item("INV_PMT_CURR") = .Item("INV_PMT")
                .Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
                .Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            Else
                Stop
            End If

            Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow

            rowARTPYMT3.Item("PYMT_BATCH_NO") = rowARTPYMT2.Item("PYMT_BATCH_NO")
            rowARTPYMT3.Item("PYMT_BATCH_LNO") = rowARTPYMT2.Item("PYMT_BATCH_LNO")
            PYMT_BATCH_ILNO += 1
            rowARTPYMT3.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
            rowARTPYMT3.Item("INV_TYPE") = .Item("INV_TYPE")
            rowARTPYMT3.Item("INV_NUM") = .Item("INV_NUM")
            rowARTPYMT3.Item("REASON_CODE") = .Item("REASON_CODE")
            rowARTPYMT3.Item("INV_DATE") = .Item("INV_DATE")
            rowARTPYMT3.Item("INV_DUE_DATE") = .Item("INV_DUE_DATE")
            rowARTPYMT3.Item("CUST_CODE_SO") = .Item("CUST_CODE_SO")
            rowARTPYMT3.Item("CUST_SHIP_TO_NO") = .Item("CUST_SHIP_TO_NO")
            rowARTPYMT3.Item("INV_CUST_PO") = .Item("INV_CUST_PO")
            rowARTPYMT3.Item("INV_BALANCE") = INV_BALANCE
            rowARTPYMT3.Item("INV_PMT") = INV_PMT
            rowARTPYMT3.Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
            rowARTPYMT3.Item("INV_WRITE_OFF") = INV_WRITE_OFF
            rowARTPYMT3.Item("INV_BALANCE_NEW") = .Item("INV_BALANCE")
            rowARTPYMT3.Item("POST_CODE") = .Item("POST_CODE")
            rowARTPYMT3.Item("SEG2_CODE") = .Item("SEG2_CODE")
            rowARTPYMT3.Item("SEG3_CODE") = .Item("SEG3_CODE")
            rowARTPYMT3.Item("SEG4_CODE") = .Item("SEG4_CODE")
            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT3.Item("INV_BALANCE_CURR") = rowARTPYMT3.Item("INV_BALANCE")
                rowARTPYMT3.Item("INV_PMT_CURR") = rowARTPYMT3.Item("INV_PMT")
                rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") = rowARTPYMT3.Item("INV_DISC_TAKEN")
                rowARTPYMT3.Item("INV_WRITE_OFF_CURR") = rowARTPYMT3.Item("INV_WRITE_OFF")
                rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") = rowARTPYMT3.Item("INV_BALANCE_NEW")
            Else
                Stop
            End If
            dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
        End With
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("SOTINVHX", YP)
        grdSOTINVHX.Text = "Entered in " & cbeYP.Text

        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            Load_SOTINVHC()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cbeYP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Private Sub grdSOTINVHX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("INV_NO").Text = e.Row.Cells("INV_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        '    UltraExplorerBar1.Groups("").Visible = False

    End Sub

    Sub Load_SOTINVHC()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Fill_Records("SOTINVHC", New String() {CUST_CODE, YP})
        grdSOTINVHC.Text = "Invoices to Credit for " & CUST_CODE & " Posted in " & cbeYP.Text
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdSOTINVHC_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHC.DoubleClickRow
        If e.Row.IsDataRow Then

            If Absx1.optFor("INV_TYPE").Value = "C" Then
                Dim ORDR_CUST_PO As String = e.Row.Cells("ORDR_CUST_PO").Value & String.Empty
                Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty

                ORDR_CUST_PO = ORDR_CUST_PO.Trim
                If ORDR_CUST_PO.Length > 0 Then
                    ASCMAIN1.sql = "Select INV_TYPE, INV_NO, ORDR_CUST_PO, INV_TOTAL_AMOUNT, INV_NO_CR" _
                        & " from SOTINVH1" _
                        & " where INV_TYPE = 'C' and ORDR_CUST_PO = :PARM1 and CUST_CODE = :PARM2"

                    Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVH1", "VV", New Object() {ORDR_CUST_PO, CUST_CODE})

                    If tbl.Rows.Count > 0 Then
                        Select Case MessageBox.Show("At least one Credit Memo has already been posted for the PO " & ORDR_CUST_PO & "." _
                                                     & vbCrLf & vbCrLf & "Do you want to view these entries?" _
                                                     & vbCrLf & vbCrLf & "Click 'Yes' to view these entries." _
                                                     & vbCrLf & "Click 'No' to not view these entries." _
                                                     & vbCrLf & "Click 'Cancel' to abort creating the Credit Memo.", "Credit Memo",
                                                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                            Case Windows.Forms.DialogResult.Cancel
                                Exit Sub

                            Case Windows.Forms.DialogResult.No

                            Case Windows.Forms.DialogResult.Yes

                                ASCMAIN1.CodeSelector.Get_SQL("")
                                ASCMAIN1.CodeSelector.VIEW_NAME = String.Empty
                                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.sql.Replace(":PARM1", "'" & ORDR_CUST_PO & "'").Replace(":PARM2", "'" & CUST_CODE & "'")
                                ASCMAIN1.CodeSelector.UseDataFromTable = tbl
                                ASCMAIN1.CodeSelector.MultipleSelections = False
                                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                                Using F As New ASFCODE1
                                    F.ShowDialog()
                                End Using

                                If MessageBox.Show("Do you want to continue creating the Credit Memo?", "Credit Memo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                                    Exit Sub
                                End If

                        End Select
                    End If

                    'Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ORDR_CUST_PO})
                    'If row IsNot Nothing Then
                    '    If MsgBox("A CR Memo has already been posted for Customer PO " & ORDR_CUST_PO _
                    '              & vbCrLf & vbCrLf & "Continue with Credit Memo?",
                    '              MsgBoxStyle.YesNo, _
                    '              "CR Memo " & row.Item("INV_NO") & " has already been posted with reference to Customer PO " & ORDR_CUST_PO) = MsgBoxResult.No Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            End If


            auto_CR = True

            Dim INV_NO As String = e.Row.Cells("INV_NO").Value
            Dim INV_TYPE As String = e.Row.Cells("INV_TYPE").Value

            rowSOTINVHC = dst.Tables("SOTINVHC").Rows.Find(New String() {INV_TYPE, INV_NO})
            Click_Command("New")
            auto_CR = False
        End If
    End Sub

    Private Sub grdSOTINVHC_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTINVHC.InitializeLayout

    End Sub

    Private Sub Print_Credit_Memo()
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Preparing Credit Memo for Printing")

            Dim REPORT_NAME As String = "SORINVP1"
            Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            If RPT = "" Then RPT = REPORT_NAME

            If Not REPORTS.ContainsKey(REPORT_NAME) Then
                REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                REPORTS(REPORT_NAME).Prepare_dst(False, "")
            End If

            Dim sql As String = " and SOTINVH1.INV_TYPE = '" & rowSOTINVH1.Item("INV_TYPE") & "' and SOTINVH1.INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'"
            Dim tempFileName As String = "Memo" & DateTime.Now.ToString("yyyyMMddHHmmss")

            REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
            Dim FILENAME As String = ""
            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "")
                Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
                If ASCMAIN1.Running_in_VS Then
                    Stop
                    .dst.WriteXml("c:\vs\aha\temp\SORINVP1.xml")
                    .dst.WriteXmlSchema("c:\vs\aha\temp\SORINVP1.xsd")
                End If

            End With

            Show_Document(FILENAME)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Credit Memo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub UltraOptionSet1_ValueChanged(sender As Object, e As EventArgs) Handles UltraOptionSet1.ValueChanged

    End Sub
End Class