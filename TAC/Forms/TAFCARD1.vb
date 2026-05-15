Public Class TAFCARD1

    Public CUST_CODE As String
    Dim CUST_NAME As String
    Dim CCPA_AMT_orig As Decimal
    Public rowARTCCPA1 As DataRow
    Public rowARTCUSTC As DataRow
    Public ShowAgedTotals As Boolean = False
    Public AGE1 As Decimal
    Public AGE2 As Decimal
    Public AGE3 As Decimal
    Public AGE4 As Decimal
    Public TOTAL_DUE As Decimal
    Public TRAN_TYPE As String
    Public AgingDescription As String
    Public AgingHeadings As New Dictionary(Of String, String)
    Public CCPA_NO_CREDIT As String
    Public CCPA_REASON As String

    Public ORDR_NO As String
    Public INV_NO As String
    Public LENS_BANK_INV_NOs As List(Of String)
    Public STMT_NO As String

    Dim test_mode As Boolean = False
    Dim rowARTCUST1 As DataRow

    Dim COLs() As String = {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_LAST4" _
    , "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE" _
    , "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
    , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE", "CUST_CREDIT_CARD_ZIP_CODE"}

    Public Sub New(ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CC_Start()


        Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
        Create_TDA(dst.Tables.Add, "ARTCCPA2", "*")
        Create_TDA(dst.Tables.Add, "ARTCUSTC", "*", 1)
 
        If Not ROWs.ContainsKey("SOTPARM1") Then
            Get_PARM("SOTPARM1")
        End If
        If Not ROWs.ContainsKey("ARTPARM1") Then
            Get_PARM("ARTPARM1")
        End If



        CCPA_AMT_orig = Val(rowARTCCPA1.Item("CCPA_AMT") & "")
        lblLeaves.Visible = False
        numLeaves.Visible = False
        optType.Value = TRAN_TYPE

        Absx1.numFor("CCPA_AMT").MaskInput = "nnn,nnn,nnn.nn"
        If CCPA_REASON = "B" Then
            Absx1.numFor("CCPA_AMT").ReadOnly = True
        End If
        grpReasonVoid.Visible = (TRAN_TYPE = "V" Or TRAN_TYPE = "C")

        If TRAN_TYPE = "V" Or TRAN_TYPE = "C" Then
            optCC.Visible = False
            Set_Read_Only(grpCCData, True)
            grpAgedTotals.Visible = False
            cmdCCSubmit.Text = optType.Items(optType.CheckedIndex).DisplayText
            'lblResponseText.Text = "Click Submit to " & optType.Items(optType.CheckedIndex).DisplayText
            lblResponseText.Appearance.ForeColor = Color.Red
            ShowAgedTotals = False
            grpReasonVoid.Text = "Reason to " & optType.Items(optType.CheckedIndex).DisplayText
        End If
        'optCC.Visible = True
        grpAgedTotals.Visible = ShowAgedTotals
        Set_Read_Only(grpAgedTotals, True)
        If ShowAgedTotals Then
            Absx1.numFor("AGE1").Value = AGE1
            Absx1.numFor("AGE2").Value = AGE2
            Absx1.numFor("AGE3").Value = AGE3
            Absx1.numFor("AGE4").Value = AGE4
            Absx1.numFor("TOTAL_DUE").Value = TOTAL_DUE

            If AgingDescription <> "" Then
                grpAgedTotals.Text = AgingDescription
            End If
            If AgingHeadings.Count <> 0 Then
                lbl1.Text = AgingHeadings("AGE1")
                lbl2.Text = AgingHeadings("AGE2")
                lbl3.Text = AgingHeadings("AGE3")
                lbl4.Text = AgingHeadings("AGE4")
            End If
        Else
            Me.Height = Me.Height - grpAgedTotals.Height
        End If

        Set_Read_Only(grpType, True)

        If ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & "" = "1" _
        Or ASCMAIN1.DBS_COMPANY.ToUpper <> "ODG" _
        Or ASCMAIN1.DBS_SERVER.ToUpper <> "ODG" Then
            test_mode = True
            lblTestMode.Visible = True
            MsgBox("Credit Card Processing is operating in Test Mode", MsgBoxStyle.OkOnly, "Please Contact ABS")
        End If

        MerchantSetup()

        If TRAN_TYPE <> "C" And TRAN_TYPE <> "V" And TRAN_TYPE <> "A" Then
            If Format(ROWs("ARTPARM1").Item("AR_PARM_CC_CUTOFF_TIME"), "HHmm") _
             < Format(Now + ASCMAIN1.NowTSD, "HHmm") And Not test_mode Then
                lblResponseText.Text = "Disabled past " & Format(ROWs("ARTPARM1").Item("AR_PARM_CC_CUTOFF_TIME"), "HH:mm")
                lblResponseText.Visible = True
                cmdCCSubmit.Enabled = False
            End If
        End If

        Me.Text &= " - " & optType.Items(optType.CheckedIndex).DisplayText
    End Sub

    Private Sub cmdCCSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCCSubmit.Click

        EMsg = ""

        If (TRAN_TYPE = "V" Or TRAN_TYPE = "C") And Absx1.txtFor("CCPA_REASON_VOID").Text = "" Then
            EMsg &= vbCr & "You Must enter a Reason for the " & optType.Items(optType.CheckedIndex).DisplayText
        End If

        Try
            Cardvalidator1.CardNumber = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text
            Cardvalidator1.CardExpMonth = Val(Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 1, 2) & "")
            Cardvalidator1.CardExpYear = Val(Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 3, 2) & "")
            Cardvalidator1.ValidateCard()
        Catch ex As Exception
            EMsg &= vbCr & ex.Message
        End Try

        If EMsg = "" Then
            If Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctMasterCard _
            Or Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctVisa _
            Or Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctVisaElectron _
            Or Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctVisaPurchase _
            Or Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctMCardPurchase _
            Or Cardvalidator1.CardType = nsoftware.IBizPtech.CardvalidatorCardTypes.vctDiscover Then
                ' WE ACCEPT THESE
            Else
                If MsgBox("You are not authorized to charge AMEX" & vbCrLf & " or any type of Credit Card other than Visa, Discover or MasterCard" & vbCrLf & " without Approval from Finance" & vbCrLf & vbCrLf & "Did you obtain approval from Finance to charge this type of card?", MsgBoxStyle.YesNo, "This type of card requires CFO Approval") = MsgBoxResult.No Then
                    EMsg = vbCr & "You must obtain approval from the Finance Dept to use this type of card"
                End If

            End If
        End If
        'Stop
        If Not Cardvalidator1.DateCheckPassed Then
            EMsg &= vbCr & "Expiration Date is Not Valid"
        End If

        If Not Cardvalidator1.DigitCheckPassed Then
            EMsg &= vbCr & "Credit Card No is not Valid"
        End If

        If Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text = "" _
        Or Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text = "" _
        Or Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text = "" Then
            EMsg &= vbCr & "Credit Card Name, Street Address & Zip Code is Mandatory"
        End If

        With Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE")
            If .Text = "" Then
                EMsg &= vbCr & "Credit Card Expiration Date is Mandatory"
            Else
                Dim EXP As String = Format$(Val(.Text), "0000")
                Dim YY As Integer = Val(Mid(ASCMAIN1.CYM, 3, 2))
                Dim Ys As String = Format$(YY, "00") & "x" & Format$(YY + 1, "00") & "x" & Format$(YY + 2, "00") & "x" & Format$(YY + 3, "00") & "x" & Format$(YY + 4, "00") & "x" & Format$(YY + 5, "00") & "x" & Format$(YY + 6, "00") & "x" & Format$(YY + 7, "00") & "x"
                Dim Ms As String = "01x" & "02x" & "03x" & "04x" & "05x" & "06x" & "07x" & "08x" & "09x" & "10x" & "11x" & "12x"
                If .Text <> EXP _
                Or InStr(Ms, Mid(EXP, 1, 2) & "x") = 0 _
                Or InStr(Ys, Mid(EXP, 3, 2) & "x") = 0 _
                Or Mid(EXP, 3, 2) & Mid(EXP, 1, 2) < Mid(ASCMAIN1.CYM, 3, 4) Then
                    EMsg &= vbCr & "Invalid Expiration Date (" & .Text & ")"
                End If
            End If
        End With

        With Absx1.txtFor("CUST_CREDIT_CARD_NO")
            If .Text = "" Then
                EMsg &= vbCr & "Credit Card No is Mandatory"
            Else
                Dim CUST_CREDIT_CARD_NO As String = Format$(CLng(.Text), "".PadLeft(Len(.Text), "0"))
                If .Text <> CUST_CREDIT_CARD_NO Then
                    EMsg &= vbCr & "Invalid Credit Card No (" & .Text & ")"
                End If
            End If
        End With

        If Val(Absx1.numFor("CCPA_AMT").Value & "") <= 0 Then
            EMsg &= vbCr & "Credit Card Amount must be > 0"
        End If

        If optCC.Value = "N" Then
            Dim CUST_CREDIT_CARD_NO As String = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text
            Dim rowARTCUSTC As DataRow = LookUp("ARTCUSTC", New String() {CUST_CODE, CUST_CREDIT_CARD_NO})
            If rowARTCUSTC IsNot Nothing Then
                EMsg &= vbCr & "Credit Card is Already on file - use Edit and Select from Lookup"
            End If
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Submit Credit Card for Processing")
            Exit Sub
        End If


        BeginTrans()

        Dim CCPA_NO As String = rowARTCCPA1.Item("CCPA_NO") & ""
        Dim CCPA_NO_orig As String = CCPA_NO

        Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD


        If CCPA_NO <> "" Then
            Dim rowARTCCPA1_orig As DataRow = dst.Tables("ARTCCPA1").NewRow
            For i As Int32 = 0 To rowARTCCPA1_orig.Table.Columns.Count - 1
                rowARTCCPA1_orig.Item(i) = rowARTCCPA1.Item(i)
            Next
            dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1_orig)
            rowARTCCPA1_orig.AcceptChanges()
            rowARTCCPA1 = dst.Tables("ARTCCPA1").Rows.Find(CCPA_NO)
        End If

        If optType.Value <> "V" Then
            If optType.Value = "C" Then
                CCPA_NO = ""
            End If
            If CCPA_NO = "" Then
                CCPA_NO = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
                rowARTCCPA1 = dst.Tables("ARTCCPA1").NewRow
                rowARTCCPA1.Item("CCPA_NO") = CCPA_NO
                rowARTCCPA1.Item("CUST_CODE") = CUST_CODE
                ' rowARTCCPA1.Item("CUST_NAME") = CUST_NAME
                rowARTCCPA1.Item("CCPA_REASON") = CCPA_REASON
                rowARTCCPA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowARTCCPA1.Item("INIT_DATE") = LAST_DATE
                rowARTCCPA1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End If

            If optType.Value = "A" Then
                rowARTCCPA1.Item("CCPA_DATE_AUTH") = LAST_DATE
            ElseIf optType.Value = "S" Then
                rowARTCCPA1.Item("CCPA_DATE_SALE") = LAST_DATE
            End If

            If optType.Value = "C" Then
                rowARTCCPA1.Item("CCPA_AMT") = -1 * Val(Absx1.numFor("CCPA_AMT").Value & "")
                rowARTCCPA1.Item("CCPA_NOTE") = Absx1.txtFor("CCPA_REASON_VOID").Text
                rowARTCCPA1.Item("CCPA_NO_CREDITED") = CCPA_NO_orig
            Else
                rowARTCCPA1.Item("CCPA_AMT") = Val(Absx1.numFor("CCPA_AMT").Value & "")
                rowARTCCPA1.Item("CCPA_NOTE") = Absx1.txtFor("CCPA_NOTE").Text
                rowARTCCPA1.Item("CCPA_REASON_VOID") = Absx1.txtFor("CCPA_REASON_VOID").Text
            End If

            LoadCCDataIntoRow(rowARTCCPA1)

            rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA1.Item("LAST_DATE") = LAST_DATE

            rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text

            If rowARTCCPA1.RowState = DataRowState.Detached Then
                dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)
            Else
                rowARTCCPA1.EndEdit()
            End If

            If optType.Value <> "C" Then
                dst.Tables("ARTCUSTC").AcceptChanges()

                If optCC.Value = "E" Then

                    rowARTCUSTC = dst.Tables("ARTCUSTC").Rows.Find(New Object() {CUST_CODE, Absx1.txtFor("CUST_CREDIT_CARD_NO").Text})
                    If rowARTCUSTC Is Nothing Then
                        Fill_Records("ARTCUSTC", CUST_CODE)
                        rowARTCUSTC = dst.Tables("ARTCUSTC").Rows.Find(New Object() {CUST_CODE, Absx1.txtFor("CUST_CREDIT_CARD_NO").Text})
                    End If

                    LoadCCDataIntoRow(rowARTCUSTC)
                    rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = "A"
                    rowARTCUSTC.Item("LAST_DATE") = LAST_DATE
                    rowARTCUSTC.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    Update_Record_TDA("ARTCUSTC")
                ElseIf optCC.Value = "N" Then
                    rowARTCUSTC = dst.Tables("ARTCUSTC").NewRow
                    rowARTCUSTC.Item("CUST_CODE") = CUST_CODE
                    LoadCCDataIntoRow(rowARTCUSTC)
                    rowARTCUSTC.Item("CUST_CREDIT_CARD_STATUS") = "A"
                    dst.Tables("ARTCUSTC").Rows.Add(rowARTCUSTC)
                    rowARTCUSTC.Item("INIT_DATE") = LAST_DATE
                    rowARTCUSTC.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    Update_Record_TDA("ARTCUSTC")
                End If
            End If

        End If


        '** the code below is replicated in individual procedures below, such as CC_Capture, CC_Sale, CC_Auth

        Try
            Dim CCPA_AMT As Decimal = Val(Absx1.numFor("CCPA_AMT").Value & "")

            Dim ZIP_CODE As String = Replace(Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text, "-", "")
            If ZIP_CODE.Length <> 5 And ZIP_CODE.Length <> 9 Then
                ZIP_CODE = ""
            End If

            'Ptcharge1.IndustryType = nsoftware.IBizPtech.PtchargeIndustryTypes.itRetail
            Ptcharge1.IndustryType = nsoftware.IBizPtech.PtchargeIndustryTypes.itDirectMarketing ' AS PER ANTHONY 09/26
            Ptcharge1.EntryDataSource = nsoftware.IBizPtech.PtchargeEntryDataSources.dsManuallyEntered

            Ptcharge1.CardNumber = Cardvalidator1.CardNumber
            Ptcharge1.CardExpMonth = Cardvalidator1.CardExpMonth
            Ptcharge1.CardExpYear = Cardvalidator1.CardExpYear

            Select Case optType.Value
                Case "A"
                    Ptcharge1.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")

                    Ptcharge1.InvoiceNumber = "1" ' REQUIRED AS PER ANTHONY
                    'Ptcharge1.GoodsIndicator = nsoftware.IBizPtech.PtchargeGoodsIndicators.giPhysicalGoods

                    Ptcharge1.Level2SalesTax = Format(CCPA_AMT * 0.08, "0.00")
                    Ptcharge1.CustomerAddress = Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
                    Ptcharge1.CustomerZip = ZIP_CODE

                    CheckTestMode()
                    Ptcharge1.AuthOnly()

                Case "S"
                    Ptcharge1.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")

                    Ptcharge1.InvoiceNumber = "1" ' REQUIRED AS PER ANTHONY

                    Ptcharge1.Level2SalesTax = Format(CCPA_AMT * 0.08, "0.00")
                    Ptcharge1.CustomerAddress = Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text

                    'Ptcharge1.GoodsIndicator = nsoftware.IBizPtech.PtchargeGoodsIndicators.giPhysicalGoods

                    Ptcharge1.CustomerZip = ZIP_CODE

                    CheckTestMode()
                    Ptcharge1.Sale()

                Case "C"
                    Ptcharge1.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")
                    'Ptcharge1.TransactionAmount = "1.00" 'credits $1.00

                    Ptcharge1.InvoiceNumber = "1" ' REQUIRED AS PER ANTHONY

                    CheckTestMode()
                    Ptcharge1.Credit()

                Case "V"
                    ASCMAIN1.sql = "Select Max (RESPONSE_RETRIEVAL_NO) from ARTCCPA2 where RESPONSE_BATCH_NO = '" & rowARTCCPA1.Item("RESPONSE_BATCH_NO") & "'"
                    Dim RESPONSE_RETRIEVAL_NO_last As String = ASCDATA1.GetDataValue

                    'Ptcharge1.TransactionAmount = Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "#.00")

                    CheckTestMode()
                    Ptcharge1.VoidTransaction(rowARTCCPA1.Item("RESPONSE_RETRIEVAL_NO"), RESPONSE_RETRIEVAL_NO_last)

                    '            cnumToVoid = component.CardNumber
                    '            refNumToVoid = component.ResponseRetrievalNumber
                    '            'All that's needed for a void:
                    '            component.CardNumber = cnumToVoid
                    '            component.VoidTransaction(refNumToVoid, "")
                    'Note that in the above void, 
                    ' if you use a new instance of the component or otherwise reset it's state, 
                    ' you must set the LastRetrievalNumber parameter as well. 
                Case Else
                    Stop
            End Select

            lblResponseText.Text = Ptcharge1.ResponseText
            If Ptcharge1.ResponseCode = "A" Then
                lblResponseText.Appearance.ForeColor = Drawing.Color.Green
            ElseIf Ptcharge1.ResponseCode = "E" Then
                lblResponseText.Appearance.ForeColor = Drawing.Color.Red
            End If

            dst.Tables("ARTCCPA2").Rows.Clear()
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = CCPA_NO
            rowARTCCPA2.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText
            rowARTCCPA2.Item("RESPONSE_SEQ_NO") = Ptcharge1.ResponseSequenceNumber
            rowARTCCPA2.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
            rowARTCCPA2.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
            rowARTCCPA2.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
            rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
            rowARTCCPA2.Item("RESPONSE_DATA") = Ptcharge1.ResponseData
            rowARTCCPA2.Item("RESPONSE_AVS") = Ptcharge1.ResponseAVS
            rowARTCCPA2.Item("RESPONSE_AUTH_SOURCE") = Ptcharge1.ResponseAuthSource
            rowARTCCPA2.Item("INIT_DATE") = LAST_DATE
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = optType.Value
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)

            If optType.Value = "V" Then
            Else
                rowARTCCPA1.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
                rowARTCCPA1.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
                rowARTCCPA1.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
                rowARTCCPA1.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
                rowARTCCPA1.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText
            End If
            If Ptcharge1.ResponseCode = "A" Then
                If optType.Value = "V" Then
                    rowARTCCPA1.Item("CCPA_STATUS") = "V"

                    If optType.Value = "V" And CCPA_REASON = "B" Then
                        For Each LENS_BANK_INV_NO As String In LENS_BANK_INV_NOs
                            ASCMAIN1.sql = "Update PPTLBKP1 set LENS_BANK_STATUS = :PARM1" _
                            & ", CCPA_NO = :PARM2" _
                            & " where LENS_BANK_INV_NO = :PARM3"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", "P", DBNull.Value, LENS_BANK_INV_NO)
                        Next
                    End If

                Else
                    If optType.Value = "A" Then
                        rowARTCCPA1.Item("CCPA_STATUS") = "T"
                    Else
                        rowARTCCPA1.Item("CCPA_STATUS") = "A"
                    End If

                    If optType.Value = "S" And CCPA_REASON = "B" Then
                        For Each LENS_BANK_INV_NO As String In LENS_BANK_INV_NOs

                            ASCMAIN1.sql = "Update PPTLBKP1 set LENS_BANK_STATUS = :PARM1" _
                            & ", CCPA_NO = :PARM2" _
                            & " where LENS_BANK_INV_NO = :PARM3"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", "A", CCPA_NO, LENS_BANK_INV_NO)

                            ASCMAIN1.sql = "Insert into ARTCCPA3 (CCPA_NO, LENS_BANK_INV_NO)" _
                            & " Values (:PARM1, :PARM2)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", CCPA_NO, LENS_BANK_INV_NO)
                        Next
                    End If

                End If
                If optType.Value = "C" Then
                    ASCMAIN1.sql = "Update ARTPYMT2 Set CCPA_NO_CREDIT = :PARM1 where CCPA_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", CCPA_NO, CCPA_NO_orig)
                    CCPA_NO_CREDIT = CCPA_NO
                End If
            Else

                ' WRITE A RECORD TO THE EVENT FOR THE SREP

                If optType.Value = "V" Then
                Else
                    rowARTCCPA1.Item("CCPA_STATUS") = "E"
                End If
            End If
            If optType.Value = "V" Then
            Else
                rowARTCCPA1.Item("ORDR_NO") = ORDR_NO
                rowARTCCPA1.Item("INV_NO") = INV_NO
                rowARTCCPA1.Item("STMT_NO") = STMT_NO
            End If
            rowARTCCPA1.Item("CCPA_TYPE") = optType.Value
            Update_Record_TDA("ARTCCPA1")
            Update_Record_TDA("ARTCCPA2")
            Dim MSG As String = "Credit Card Payment Submitted"
            If optType.Value = "A" Then
                MSG = "Credit Card Authorized for " & Format(Val(Absx1.numFor("CCPA_AMT").Value & ""), "$#,###.00")
            End If
            If optType.Value = "V" Then
                MSG = "Credit Card Payment Voided"
            ElseIf optType.Value = "C" Then
                MSG = "CC Credit Processed"
            End If
            If Ptcharge1.ResponseCode <> "A" Then
                MSG &= " - With Error"
            End If
            CommitTrans(MSG)

        Catch ex As Exception
            Rollback("Please call ABS - Error processing Card Payment: " & ex.Message)

        End Try


        Me.Close()

    End Sub

    Sub LoadCCDataIntoRow(ByVal row As DataRow)
        For Each COL As String In COLs
            If row.Item(COL) & "" <> Absx1.txtFor(COL).Text Then
                row.Item(COL) = Absx1.txtFor(COL).Text
            End If
        Next
    End Sub

    Sub CC_Clear()
        For Each COL As String In COLs
            If COL <> "CUST_CREDIT_CARD_LAST4" Then
                Absx1.txtFor(COL).Text = ""
            End If
        Next
        Call CC_EnableControls(False)
    End Sub

    Sub CC_Start()

        CUST_CODE = rowARTCCPA1.Item("CUST_CODE")
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        CUST_NAME = rowARTCUST1.Item("CUST_NAME") & ""
        grpCCData.Text = "Customer " & CUST_CODE
        Absx1.txtFor("CUST_NAME").Text = CUST_NAME
        Absx1.txtFor("CUST_CONTACT").Text = rowARTCUST1.Item("CUST_CONTACT") & ""
        Absx1.medFor("CUST_PHONE").Value = rowARTCUST1.Item("CUST_PHONE") & ""
        lblResponseText.Text = ""
        optCC.Value = "X"

        ASCMAIN1.sql = "Select Sum (CUST_PYMT_AMT) from ARTPYMT1,ARTPYMT2 where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO and ARTPYMT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and CUST_CODE = '" & CUST_CODE & "' and NVL(PYMT_DELETED,'0') <> '1'"
        Dim PAID_MTD As Decimal = Val(ASCDATA1.GetDataValue & "")
        Absx1.numFor("PAID_MTD").Value = PAID_MTD
    End Sub

    Sub CC_EnableControls(ByVal tf As Boolean)
        Absx1.txtFor("CUST_CREDIT_CARD_NO").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_NAME").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_CITY").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_STATE").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").ReadOnly = Not tf
        Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").ReadOnly = Not tf

        Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = "*"
        cmdUseCustomerAddress.Visible = tf
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CCPA_AMT"
                If CCPA_AMT_orig <> 0 Then
                    Dim CCPA_AMT = Val(Absx1.numFor("CCPA_AMT").Value & "")
                    If CCPA_AMT_orig - CCPA_AMT <= 0 Then
                        numLeaves.Value = 0
                    Else
                        numLeaves.Value = CCPA_AMT_orig - CCPA_AMT
                    End If
                    ' numLeaves.Value = CCPA_AMT_orig - CCPA_AMT
                    If CCPA_AMT <> CCPA_AMT_orig Then
                        lblLeaves.Visible = True
                        numLeaves.Visible = True
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_CARD_NO"
                Dim txtctl As UltraWinEditors.UltraTextEditor
                txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                If txtctl.ReadOnly Then
                    Exit Sub
                End If
                If txtctl.PasswordChar = "*" Then
                    e.Handled = True
                    e.SuppressKeyPress = True
                End If
        End Select

        'If e.KeyCode = Windows.Forms.Keys.Enter Then

        'End If
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CREDIT_CARD_NO"
                Dim CUST_CREDIT_CARD_NO As String = Absx1.txtFor("CUST_CREDIT_CARD_NO").Text
                If Len(CUST_CREDIT_CARD_NO) > 4 Then
                    Absx1.txtFor("CUST_CREDIT_CARD_LAST4").Text = Mid(CUST_CREDIT_CARD_NO, Len(CUST_CREDIT_CARD_NO) - 3, 4)
                Else
                    Absx1.txtFor("CUST_CREDIT_CARD_LAST4").Text = ""
                End If
                Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = ""

                If CUST_CREDIT_CARD_NO.Length >= 15 Then
                    Try
                        Cardvalidator1.CardNumber = CUST_CREDIT_CARD_NO
                        'Cardvalidator1.CardExpMonth = Val(Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 1, 2) & "")
                        'Cardvalidator1.CardExpYear = Val(Mid$(Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text, 3, 2) & "")
                        Cardvalidator1.ValidateCard()

                    Catch ex As Exception

                    End Try

                    Dim IMAGE_FILE As String = ""
                    Select Case Cardvalidator1.CardType
                        Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctAmex
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "AMEX"
                            IMAGE_FILE = "AMEX.GIF"
                        Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctMasterCard
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "MSTR"
                            IMAGE_FILE = "MSTR.GIF"
                        Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctDiscover
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "DISC"
                            IMAGE_FILE = "DISC.GIF"
                        Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctVisa
                            Absx1.txtFor("CUST_CREDIT_CARD_TYPE").Text = "VISA"
                            IMAGE_FILE = "VISA.GIF"
                        Case Else

                    End Select
                    If IMAGE_FILE <> "" Then
                        picCC.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "\ABS\CC\", IMAGE_FILE)
                    Else
                        picCC.Image = Nothing
                    End If
                Else
                    picCC.Image = Nothing
                End If

        End Select

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CREDIT_CARD_NO"
                sql_where = "CUST_CODE = '" & CUST_CODE & "'"
                If optCC.Value = "N" Then
                    Cancel = True
                End If
            Case "CUST_CREDIT_CARD_LAST4"
                sql_where = "CUST_CODE = '" & CUST_CODE & "'"

        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CREDIT_CARD_NO"
                Dim rowARTCUSTC As DataRow = LookUp("ARTCUSTC", New String() {CUST_CODE, txtctl.Text})
                LoadDataFromARTCUSTC(rowARTCUSTC)
                'Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = "*"
                'CC_EnableControls(False)
        End Select
    End Sub

    Sub LoadDataFromARTCUSTC(ByVal rowARTCUSTC As DataRow)
        With rowARTCUSTC
            Absx1.txtFor("CUST_CREDIT_CARD_NAME").Text = .Item("CUST_CREDIT_CARD_NAME") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text = .Item("CUST_CREDIT_CARD_ADDR1") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_CITY").Text = .Item("CUST_CREDIT_CARD_CITY") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_STATE").Text = .Item("CUST_CREDIT_CARD_STATE") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_ZIP_CODE").Text = .Item("CUST_CREDIT_CARD_ZIP_CODE") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Text = .Item("CUST_CREDIT_CARD_EXP_DATE") & ""
            Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text = .Item("CUST_CREDIT_CARD_VER_CODE") & ""
        End With
    End Sub

    Private Sub MerchantSetup()
        Ptcharge1.Server = ROWs("SOTPARM1").Item("SO_PARM_CC_SERVER")
        Ptcharge1.MerchantNumber = ROWs("SOTPARM1").Item("SO_PARM_CC_MERCHANT_NO")
        Ptcharge1.TerminalNumber = ROWs("SOTPARM1").Item("SO_PARM_CC_TERMINAL_NO")
        Ptcharge1.ClientNumber = ROWs("SOTPARM1").Item("SO_PARM_CC_CLIENT_NO")
        Ptcharge1.UserId = ROWs("SOTPARM1").Item("SO_PARM_CC_USER_ID")
        Ptcharge1.Password = ROWs("SOTPARM1").Item("SO_PARM_CC_PASSWORD")
        'Stop
        'If ASCMAIN1.DBS_SERVER = "ODG" And ASCMAIN1.DBS_COMPANY = "ODG" Then
        '    Stop
        'Else
        '    Ptcharge1.Server = ROWs("SOTPARM1").Item("SO_PARM_CC_SERVER_TEST")
        '    'This is a test server URL
        '    Ptcharge1.Server = "https://netconnectvar.paymentech.net/NetConnect/controller"
        '    Ptcharge1.MerchantNumber = "700000000125"
        '    Ptcharge1.TerminalNumber = "100"
        '    Ptcharge1.ClientNumber = "0002"
        '    Ptcharge1.UserId = "nsoftware01"
        '    Ptcharge1.Password = "nsoftwarepw01"
        'End If
    End Sub

    Private Sub optCC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCC.ValueChanged

        Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = "*"
        Absx1.txtFor("CUST_CREDIT_CARD_NO").Appearance.BackColor = Drawing.Color.Empty

        Select Case optCC.Value
            Case "N"
                Me.CC_Clear()
                Call CC_EnableControls(True)
                Absx1.txtFor("CUST_CREDIT_CARD_NO").PasswordChar = ""
                Absx1.txtFor("CUST_CREDIT_CARD_NO").Appearance.BackColor = Drawing.Color.AliceBlue
                Absx1.txtFor("CUST_CREDIT_CARD_NO").Focus()
            Case "E"
                LoadCCDataFromInitialCCPA()
                Call CC_EnableControls(True)
                Absx1.txtFor("CUST_CREDIT_CARD_EXP_DATE").Focus()
            Case "X"
                LoadCCDataFromInitialCCPA()
        End Select
    End Sub

    Sub LoadCCDataFromInitialCCPA()
        For Each COLUMN_NAME As String In New String() _
        {"CUST_CREDIT_CARD_NO" _
        , "CUST_CREDIT_CARD_NAME", "CUST_CREDIT_CARD_ADDR1" _
        , "CUST_CREDIT_CARD_CITY", "CUST_CREDIT_CARD_STATE" _
        , "CUST_CREDIT_CARD_ZIP_CODE" _
        , "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE" _
        , "CCPA_AMT", "CCPA_NOTE", "CCPA_REASON_VOID"}
            Absx1.CtlFor(COLUMN_NAME).Text = rowARTCCPA1.Item(COLUMN_NAME) & ""
        Next

        Call CC_EnableControls(False)
    End Sub

    Sub CheckTestMode()
        If test_mode Or optCC.Value = "A" Then
            Ptcharge1.CardNumber = "371055358751001"
            Ptcharge1.TransactionAmount = 0.01
        End If
    End Sub

    Private Sub cmdUseCustomerAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUseCustomerAddress.Click
        For Each COL As String In New String() {"ADDR1", "CITY", "STATE", "ZIP_CODE"}
            Absx1.txtFor("CUST_CREDIT_CARD_" & COL).Text = rowARTCUST1.Item("CUST_" & COL) & ""
        Next
    End Sub

    Public Function CC_Capture(ByVal CCPA_AMT As Decimal) As String

        ' CC_Capture, CC_Authorize & CC_Sale are called by SH modules which instantiate this form as a class
        ' when this happens, Form Load does not fire 
        '  (and it is good that it does not because there is a lot of UI going on in Form Load)
        ' A notable difference between CC_Capture and the others is that CC_Capture does not update Oracle,
        '  whereas CC_Authorize and CC_Sale do
        ' However, the datalayer that is used needs to (still) be the calling form's datalayer,
        '  since Form Load is not called - which is where we Create_TDA
        ' So these routines will continue to use the calling form's datalayer 
        '  until we come up with a more elegant solution

        Initialize_DataLayer() ' Instantiation of this form does not call Form_Load

        'rowARTCCPA1 = dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1.ItemArray)
        'rowARTCCPA1.AcceptChanges()

        MerchantSetup()
        Prepare_Component(CCPA_AMT)
        Ptcharge1.Capture(rowARTCCPA1.Item("RESPONSE_APPROVAL_CODE"))

        Dim CCPA_NO_CAPTURE As String = ""
        If Ptcharge1.ResponseCode = "A" Then
            rowARTCCPA1.Item("CCPA_STATUS") = "C"
            CCPA_NO_CAPTURE = Create_CC_Capture_Entry(rowARTCCPA1, "")
        Else
            rowARTCCPA1.Item("CCPA_STATUS") = "B"

            ' Record Error trying to Capture Previously Authorized Sale
            Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
            rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
            rowARTCCPA2.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText
            rowARTCCPA2.Item("RESPONSE_SEQ_NO") = Ptcharge1.ResponseSequenceNumber
            rowARTCCPA2.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
            rowARTCCPA2.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
            rowARTCCPA2.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
            rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
            rowARTCCPA2.Item("RESPONSE_DATA") = Ptcharge1.ResponseData
            rowARTCCPA2.Item("RESPONSE_AVS") = Ptcharge1.ResponseAVS
            rowARTCCPA2.Item("RESPONSE_AUTH_SOURCE") = Ptcharge1.ResponseAuthSource
            rowARTCCPA2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            rowARTCCPA2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowARTCCPA2.Item("CCPA_TYPE") = "E"
            dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)

        End If

        'Update_Record_TDA("ARTCCPA1")
        'Update_Record_TDA("ARTCCPA2")

        Return CCPA_NO_CAPTURE

    End Function

    Public Function Create_CC_Capture_Entry(ByRef rowARTCCPA1_AUTH As DataRow, ByVal EntryNote As String) As String

        Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD

        Dim CCPA_NO_CAPTURED As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
        Dim rowARTCCPA1_Capture As DataRow = dst.Tables("ARTCCPA1").NewRow
        rowARTCCPA1_Capture.Item("CCPA_NO") = CCPA_NO_CAPTURED

        rowARTCCPA1_AUTH.Item("CCPA_NO_CAPTURE") = CCPA_NO_CAPTURED
        rowARTCCPA1_AUTH.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_AUTH.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CCPA_NO_AUTH") = rowARTCCPA1_AUTH.Item("CCPA_NO")

        rowARTCCPA1_Capture.Item("CUST_CODE") = rowARTCCPA1_AUTH.Item("CUST_CODE")
        rowARTCCPA1_Capture.Item("CCPA_STATUS") = Ptcharge1.ResponseCode
        rowARTCCPA1_Capture.Item("CCPA_REASON") = "C"
        rowARTCCPA1_Capture.Item("CCPA_NOTE") = EntryNote
        rowARTCCPA1_Capture.Item("CCPA_AMT") = Ptcharge1.TransactionAmount
        rowARTCCPA1_Capture.Item("CCPA_DATE_AUTH") = rowARTCCPA1_AUTH.Item("CCPA_DATE_AUTH")
        rowARTCCPA1_Capture.Item("CCPA_DATE_SALE") = LAST_DATE
        'rowARTCCPA1.Item("CCPA_AUTH") = ""
        rowARTCCPA1_Capture.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("INIT_DATE") = LAST_DATE
        rowARTCCPA1_Capture.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1_Capture.Item("LAST_DATE") = LAST_DATE

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NO") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_VER_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_NAME") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ADDR1")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_CITY") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_CITY")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_STATE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_STATE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ZIP_CODE")
        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_LAST4") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_LAST4")

        rowARTCCPA1_Capture.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
        rowARTCCPA1_Capture.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
        rowARTCCPA1_Capture.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
        rowARTCCPA1_Capture.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
        rowARTCCPA1_Capture.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText

        rowARTCCPA1_Capture.Item("CCPA_NOTE") = "Pre-Auth Sale"
        ' this needs to go into a separate column

        rowARTCCPA1_Capture.Item("CCPA_TYPE") = "S"
        rowARTCCPA1_Capture.Item("ORDR_NO") = rowARTCCPA1_AUTH.Item("ORDR_NO")
        rowARTCCPA1_Capture.Item("INV_NO") = rowARTCCPA1_AUTH.Item("INV_NO")

        rowARTCCPA1_Capture.Item("CUST_CREDIT_CARD_TYPE") = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_TYPE")
        rowARTCCPA1_Capture.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1_Capture)


        If rowARTCCPA1_AUTH.Item("ORDR_NO") <> "" Then
            'SOCMAIN1.Record_Event_SOTORDRE(rowARTCCPA1_AUTH.Item("ORDR_NO"), _
            '                      rowARTCCPA1_AUTH.Item("LAST_DATE"), _
            '                      rowARTCCPA1_AUTH.Item("LAST_OPER"), _
            '                      "", "Sale Captured Inv " & rowARTCCPA1_AUTH.Item("INV_NO") & " " & Format(Val(rowARTCCPA1_Capture.Item("CCPA_AMT") & ""), "$#,##0.00"))
        End If


        Record_Response2(rowARTCCPA1_Capture)

        Return CCPA_NO_CAPTURED
    End Function

    Sub Initalize_CCPA()

        Initialize_DataLayer()
        MerchantSetup()

        Dim CCPA_AMT As Decimal = Val(rowARTCCPA1.Item("CCPA_AMT") & "")
        Dim CCPA_NO As String = ASCMAIN1.Next_Control_No("ARTCCPA1.CCPA_NO")
        Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD

        rowARTCCPA1.Item("CCPA_NO") = CCPA_NO
        rowARTCCPA1.Item("ORDR_NO") = ORDR_NO
        rowARTCCPA1.Item("CCPA_TYPE") = TRAN_TYPE
        rowARTCCPA1.Item("INIT_DATE") = INIT_DATE
        rowARTCCPA1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1.Item("LAST_DATE") = INIT_DATE
        rowARTCCPA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTCCPA1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        Dim CUST_CREDIT_CARD_NO As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO")
        Dim MMYY As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE")
        rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4") = Mid(CUST_CREDIT_CARD_NO, Len(CUST_CREDIT_CARD_NO) - 3, 4)

        Try
            Cardvalidator1.CardNumber = CUST_CREDIT_CARD_NO
            Cardvalidator1.CardExpMonth = Mid(MMYY, 1, 2)
            Cardvalidator1.CardExpYear = Mid(MMYY, 3, 2)
            Cardvalidator1.ValidateCard()
        Catch ex As Exception
            'Stop
        End Try

        Dim CUST_CREDIT_CARD_TYPE As String = ""
        Select Case Cardvalidator1.CardType
            Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctAmex
                CUST_CREDIT_CARD_TYPE = "AMEX"
            Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctMasterCard
                CUST_CREDIT_CARD_TYPE = "MSTR"
            Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctDiscover
                CUST_CREDIT_CARD_TYPE = "DISC"
            Case nsoftware.IBizPtech.CardvalidatorCardTypes.vctVisa
                CUST_CREDIT_CARD_TYPE = "VISA"
            Case Else
        End Select
        rowARTCCPA1.Item("CUST_CREDIT_CARD_TYPE") = CUST_CREDIT_CARD_TYPE

        dst.Tables("ARTCCPA1").Rows.Add(rowARTCCPA1)

        Prepare_Component(CCPA_AMT)
    End Sub

    Sub Prepare_Component(ByVal CCPA_AMT As Decimal)

        Ptcharge1.CardNumber = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO")
        Dim MMYY As String = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE")
        Ptcharge1.CardExpMonth = Mid(MMYY, 1, 2)
        Ptcharge1.CardExpYear = Mid(MMYY, 3, 2)
        Ptcharge1.TransactionAmount = Format(CCPA_AMT, "#.00")
        Ptcharge1.IndustryType = nsoftware.IBizPtech.PtchargeIndustryTypes.itDirectMarketing ' AS PER ANTHONY 09/26
        Ptcharge1.EntryDataSource = nsoftware.IBizPtech.PtchargeEntryDataSources.dsManuallyEntered

        Ptcharge1.InvoiceNumber = "1" ' REQUIRED AS PER ANTHONY
        Ptcharge1.Level2SalesTax = Format(CCPA_AMT * 0.08, "0.00")
        Ptcharge1.CustomerAddress = rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") & ""

        Dim ZIP_CODE As String = Replace(rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") & "", "-", "")
        If ZIP_CODE.Length <> 5 And ZIP_CODE.Length <> 9 Then
            ZIP_CODE = ""
        End If
        Ptcharge1.CustomerZip = ZIP_CODE
    End Sub

    Public Function CC_Authorize(ByVal BeginCommit) As String

        ' see note at CC_Capture

        Initalize_CCPA()
        Ptcharge1.AuthOnly()
        Record_Response()
        Record_Audit(BeginCommit)

        Return Ptcharge1.ResponseCode

    End Function

    Public Function CC_Sale(ByVal BeginCommit) As String

        ' see note at CC_Capture

        Initalize_CCPA()
        Ptcharge1.Sale()
        Record_Response()
        Record_Audit(BeginCommit)

        Return Ptcharge1.ResponseCode

    End Function

    Sub Record_Response()
        rowARTCCPA1.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
        rowARTCCPA1.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
        rowARTCCPA1.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
        rowARTCCPA1.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
        rowARTCCPA1.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText

        Record_Response2(rowARTCCPA1)
    End Sub

    Sub Record_Response2(ByVal rowARTCCPA1 As DataRow)
        Dim rowARTCCPA2 As DataRow = dst.Tables("ARTCCPA2").NewRow
        rowARTCCPA2.Item("CCPA_NO") = rowARTCCPA1.Item("CCPA_NO")
        rowARTCCPA2.Item("RESPONSE_TEXT") = Ptcharge1.ResponseText
        rowARTCCPA2.Item("RESPONSE_SEQ_NO") = Ptcharge1.ResponseSequenceNumber
        rowARTCCPA2.Item("RESPONSE_RETRIEVAL_NO") = Ptcharge1.ResponseRetrievalNumber
        rowARTCCPA2.Item("RESPONSE_CODE") = Ptcharge1.ResponseCode
        rowARTCCPA2.Item("RESPONSE_BATCH_NO") = Ptcharge1.ResponseBatchNumber
        rowARTCCPA2.Item("RESPONSE_APPROVAL_CODE") = Ptcharge1.ResponseApprovalCode
        rowARTCCPA2.Item("RESPONSE_DATA") = Ptcharge1.ResponseData
        rowARTCCPA2.Item("RESPONSE_AVS") = Ptcharge1.ResponseAVS
        rowARTCCPA2.Item("RESPONSE_AUTH_SOURCE") = Ptcharge1.ResponseAuthSource
        rowARTCCPA2.Item("INIT_DATE") = rowARTCCPA1.Item("INIT_DATE")
        rowARTCCPA2.Item("INIT_OPER") = rowARTCCPA1.Item("INIT_OPER")
        rowARTCCPA2.Item("CCPA_TYPE") = rowARTCCPA1.Item("CCPA_TYPE")
        dst.Tables("ARTCCPA2").Rows.Add(rowARTCCPA2)
    End Sub

    Sub Record_Audit(ByVal BeginCommit)

        If BeginCommit Then
            BeginTrans()
        End If

        Dim EVENT_DESC As String = ""
        If Ptcharge1.ResponseCode = "A" Then

            If rowARTCCPA1.Item("CCPA_TYPE") = "S" Then
                rowARTCCPA1.Item("CCPA_DATE_SALE") = rowARTCCPA1.Item("LAST_DATE")
                rowARTCCPA1.Item("CCPA_STATUS") = "A"
                EVENT_DESC = "CC Sale Appr "
            ElseIf rowARTCCPA1.Item("CCPA_TYPE") = "A" Then
                rowARTCCPA1.Item("CCPA_DATE_AUTH") = rowARTCCPA1.Item("LAST_DATE")
                rowARTCCPA1.Item("CCPA_STATUS") = "T"
                EVENT_DESC = "CC Auth Appr "
            End If

        Else
            If rowARTCCPA1.Item("CCPA_TYPE") = "S" Then
                rowARTCCPA1.Item("CCPA_STATUS") = "E"
                EVENT_DESC = "CC Sale Decl "
            ElseIf rowARTCCPA1.Item("CCPA_TYPE") = "A" Then
                rowARTCCPA1.Item("CCPA_STATUS") = "E"
                EVENT_DESC = "CC Auth Decl "
            End If
        End If

        If ORDR_NO <> "" Then
            'SOCMAIN1.Record_Event_SOTORDRE(ORDR_NO, _
            '                      rowARTCCPA1.Item("LAST_DATE"), _
            '                      rowARTCCPA1.Item("LAST_OPER"), _
            '                      "", EVENT_DESC & Format(Val(rowARTCCPA1.Item("CCPA_AMT") & ""), "$#,##0.00"))
        End If

        frmASFBASE1.Update_Record_TDA("ARTCCPA1")
        frmASFBASE1.Update_Record_TDA("ARTCCPA2")

        If BeginCommit Then
            CommitTrans()
        End If

    End Sub

    Sub Initialize_DataLayer()

        ' this will point dst and ROWs back to the calling form's datalayer

        ' This is necessary because:
        ' 1) sometimes the calling form needs to look at the data created in ARTCCPA1 (see SHFORDRX, which updates oracle in its own update routines)
        ' 2) sometimed the calling form handles the update to oracle (see Web Order import, which checks for Auth Declines)

        ROWs = frmASFBASE1.ROWs
        dst = frmASFBASE1.dst
    End Sub
End Class