Public Class SPFDPMT1

    Dim rowSPTDPMT1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim PYMT_NO As String
    Dim PYMT_BATCH_NO As String
    Dim sqlARTOPENX As String
    Dim ACC_CTL_NOs As New List(Of String)
    Dim auto_writeoff As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")

        With dst
            ASCMAIN1.sql = "Select * from SPTDCOMC where PYMT_NO is Null"
            Create_TDA(.Tables.Add, "SPTDCOMX", "**", 0, False)
            With .Tables("SPTDCOMX")
                .Columns.Add("AMT_TO_WOFF", GetType(System.Decimal))
                .Columns.Add("AMT_COMM_NEW", GetType(System.Decimal), "ISNULL(AMT_COMM,0)-ISNULL(AMT_TO_WOFF,0)")
                .Columns.Add("WRITE_OFF")
                .Columns("WRITE_OFF").DefaultValue = "0"
            End With

            sqlARTOPENX = "Select ARTOPEN1.CUST_CODE, ARTOPEN1.INV_TYPE, ARTOPEN1.INV_NUM" & vbCrLf _
                & ", ARTOPEN1.INV_DATE, ARTOPEN1.INV_CUST_PO, ARTOPEN1.INV_TOTAL_AMOUNT, ARTOPEN1.INV_BALANCE" & vbCrLf _
                & ", ARTOPEN1.INIT_OPER, ARTOPEN1.INIT_DATE, ARTOPEN1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ARTOPEN1.INV_NOTES, ARTOPEN1.ORDR_TYPE_CODE, ARTOPEN1.OPS_YYYYPP" & vbCrLf
            ASCMAIN1.sql = sqlARTOPENX _
                & " from ARTOPEN1 where INV_TYPE = 'B' and REASON_CODE = :PARM1 and CUST_CODE = NVL(:PARM2,CUST_CODE)"
            Create_TDA(.Tables.Add, "ARTOPENX", "**", 0, False, "VV")
            With .Tables("ARTOPENX")
                .Columns.Add("AMT_TO_APPLY", GetType(System.Decimal))
                .Columns.Add("INV_BALANCE_NEW", GetType(System.Decimal), "ISNULL(INV_BALANCE,0)-ISNULL(AMT_TO_APPLY,0)")
            End With

            'ASCMAIN1.sql = "Select ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, ARTPYMT2.CUST_CODE, ARTPYMT2.CUST_NAME, ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE, ARTPYMT2.PYMT_NOTE, ARTPYMT2.INIT_DATE, ARTPYMT2.INIT_OPER from ARTPYMT2,ARTPYMT1" & vbCrLf _
            '    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            '    & "   and ARTPYMT1.PYMT_SOURCE = 'D'"
            ASCMAIN1.sql = "Select * from SPTDPMT1"
            Create_TDA(.Tables.Add, "SPTDPMTX", "**", 0, False)

            ASCMAIN1.sql = "Select OPS_YYYYPP, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", COUNT (*) COLLECTIONS" & vbCrLf _
                & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
                & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
                & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
                & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
                & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
                & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
                & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
                & " from SPTDCOMB" & vbCrLf _
                & " where OPS_YYYYPP = :PARM1 and CUST_CODE = :PARM2 and HC_CODE = :PARM3" & vbCrLf _
                & " group by OPS_YYYYPP, CUST_CODE, CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTDCOMS", "**", 0, False, "VVV", 3)

            ASCMAIN1.sql = "Select OPS_YYYYPP, CUST_CODE, COLLECTION_CODE" & vbCrLf _
                & ", COUNT (*) STORES" & vbCrLf _
                & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
                & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
                & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
                & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
                & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
                & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
                & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
                & " from SPTDCOMB" & vbCrLf _
                & " where OPS_YYYYPP = :PARM1 and CUST_CODE = :PARM2 and HC_CODE = :PARM3" & vbCrLf _
                & " group by OPS_YYYYPP, CUST_CODE, COLLECTION_CODE"
            Create_TDA(.Tables.Add, "SPTDCOML", "**", 0, False, "VVV", 3)

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.TRADE_CLASS_CODE, SPTDCOM1.DEMO_COMM_PCT" & vbCrLf _
                & " from ARTCUST1,SPTDCOM1 where SPTDCOM1.CUST_CODE (+) = ARTCUST1.CUST_CODE and ARTCUST1.CUST_CODE =:PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SPTDCOMC where CUST_CODE = :PARM1 and OPS_YYYYPP_PAID is Null"
            Create_TDA(.Tables.Add, "SPTDCOMC", "**", 0, True, "V")
            With .Tables("SPTDCOMC")
                .Columns.Add("APPLY")
                .Columns("APPLY").DefaultValue = "0"
                .Columns.Add("AMT_COMM_ADD_CALC", GetType(System.Decimal), "IIF(APPLY='0' OR ISNULL(AMT_COMM,0)=ISNULL(AMT_COMM_PAID,0) OR ISNULL(AMT_COMM,0)>ISNULL(AMT_COMM_PAID,0),NULL,ISNULL(AMT_COMM_PAID,0)-ISNULL(AMT_COMM,0))")
                .Columns.Add("AMT_COMM_RED_CALC", GetType(System.Decimal), "IIF(APPLY='0' OR ISNULL(AMT_COMM,0)=ISNULL(AMT_COMM_PAID,0) OR ISNULL(AMT_COMM,0)<ISNULL(AMT_COMM_PAID,0),NULL,ISNULL(AMT_COMM,0)-ISNULL(AMT_COMM_PAID,0))")
                .Columns.Add("LEAVE_OPEN")
                .Columns("LEAVE_OPEN").DefaultValue = "0"
                .Columns.Add("AMT_COMM_ADD", GetType(System.Decimal), "IIF(LEAVE_OPEN='1',NULL,ISNULL(AMT_COMM_ADD_CALC,0))")
                .Columns.Add("AMT_COMM_RED", GetType(System.Decimal), "IIF(LEAVE_OPEN='1',NULL,ISNULL(AMT_COMM_RED_CALC,0))")
                .Columns.Add("BALANCE", GetType(System.Decimal), "IIF(LEAVE_OPEN='0',NULL,ISNULL(AMT_COMM_RED_CALC,0)-ISNULL(AMT_COMM_ADD_CALC,0))")
            End With

            ASCMAIN1.sql = "Select * from SPTDCOMB where CUST_CODE = :PARM1 and OPS_YYYYPP = :PARM2"
            Create_TDA(.Tables.Add, "SPTDCOMB", "**", 0, True, "VV")

            ASCMAIN1.sql = "Select * from SPTDCOMC where CUST_CODE = :PARM1 and OPS_YYYYPP_PAID is Not Null"
            Create_TDA(.Tables.Add, "SPTDCOMP", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SPTDPMT1", "*")
            Create_TDA(.Tables.Add, "SPTDPMT2", "*", 1)

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")

            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
            .Tables("ARTPYMT3").Columns.Add("SELECTED")
            .Tables("ARTPYMT3").Columns("SELECTED").DefaultValue = "0"
            .Tables("ARTPYMT3").Columns.Add("CUST_CODE")

        End With

        grdSPTDCOMX.DataSource = dst.Tables("SPTDCOMX")
        grdSPTDCOMC.DataSource = dst.Tables("SPTDCOMC")
        grdSPTDCOMP.DataSource = dst.Tables("SPTDCOMP")

        grdARTOPENX.DataSource = dst.Tables("ARTOPENX")
        grdSPTDPMTX.DataSource = dst.Tables("SPTDPMTX")

        grdSPTDCOMS.DataSource = dst.Tables("SPTDCOMS")
        grdSPTDCOML.DataSource = dst.Tables("SPTDCOML")


        grd_Appearance_LightGray(grdSPTDCOMX)
        grd_Appearance_LightGray(grdSPTDCOMC)
        grd_Appearance_LightGray(grdSPTDCOMP)
        grd_Appearance_LightGray(grdARTOPENX)
        grd_Appearance_LightGray(grdSPTDPMTX)
        grd_Appearance_LightGray(grdSPTDCOMS)
        grd_Appearance_LightGray(grdSPTDCOML)

        Create_Summary(grdSPTDCOMX, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTDCOMX, New String() {"AMT_COMM", "AMT_TO_WOFF", "AMT_COMM_NEW", "WRITE_OFF"})

        Create_Summary(grdSPTDCOMC, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTDCOMC, New String() {"QTY_SOLD", "AMT_SOLD", "AMT_COMM", "AMT_SOLD_CLAIMED", "AMT_COMM_PAID", "APPLY", "AMT_COMM_ADD", "AMT_COMM_RED", "BALANCE"})

        Create_Summary(grdSPTDCOMP, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTDCOMP, New String() {"QTY_SOLD", "AMT_SOLD", "AMT_COMM", "AMT_SOLD_CLAIMED", "AMT_COMM_PAID"})

        Create_Summary(grdARTOPENX, "INV_NUM", "Count")
        Create_Summary(grdARTOPENX, New String() {"INV_BALANCE", "AMT_TO_APPLY", "INV_BALANCE_NEW"})

        Create_Summary(grdSPTDCOMS, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTDCOMS, New String() {"QTY_SOLD", "AMT_SOLD", "QTY_EOW", "AMT_EOW", "AMT_COMM"})

        Create_Summary(grdSPTDCOML, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTDCOML, New String() {"QTY_SOLD", "AMT_SOLD", "QTY_EOW", "AMT_EOW", "AMT_COMM"})

        Create_Summary(grdSPTDPMTX, "PYMT_NO", "Count")
        Create_Summary(grdSPTDPMTX, New String() {"PYMT_REF_AMT", "DEMO_COMM_OFFSET", "DEMO_COMM_CREDITED", "DEMO_COMM_APPLIED", "DEMO_COMM_ADD_EXP"})

        Show_Filter(grdSPTDCOMX, True)
        grdSPTDCOMX.DisplayLayout.GroupByBox.Hidden = False

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTDCOMX, grdSPTDCOMC, grdSPTDCOMP, grdSPTDCOMS, grdSPTDCOML}
            With grd.DisplayLayout.Bands(0)
                .Columns("QTY_SOLD").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns("AMT_SOLD").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns("AMT_COMM").Header.Appearance.BackColor2 = Drawing.Color.Gold
                If grd.Name = "grdSPTDCOMS" Or grd.Name = "grdSPTDCOML" Then
                    .Columns("QTY_EOW").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    .Columns("AMT_EOW").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If

                If grd.Name = "grdSPTDCOMX" Then
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    .Columns("WRITE_OFF").Header.Appearance.BackColor2 = Drawing.Color.Orange
                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        If gcol.Key = "WRITE_OFF" Then '       If gcol.Key = "AMT_TO_WOFF" Then
                            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        Else
                            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    Next
                End If
            End With
        Next

        With grdARTOPENX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"PAY", "AMT_TO_APPLY", "INV_BALANCE_NEW"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    If gcol.Key = "AMT_TO_APPLY" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End If
            Next
        End With

        With grdSPTDCOMC.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"APPLY", "LEAVE_OPEN", "BALANCE", "AMT_SOLD_CLAIMED", "AMT_COMM_PAID", "AMT_COMM_ADD", "AMT_COMM_RED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    If gcol.Key = "AMT_SOLD_CLAIMED" Or gcol.Key = "AMT_COMM_PAID" Or gcol.Key = "APPLY" Or gcol.Key = "LEAVE_OPEN" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    End If
                End If
            Next
        End With

        Set_Read_Only(grpTotals, True)

        Bind_Controls(grpARTCUST1, "ARTCUST1")

        grpStores.Visible = False
        grpARTCUST1.Visible = False

        If ASCMAIN1.Running_in_VS Then
            btnWriteOff.Visible = True
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Validate_Code("CUST_CODE")

            Case "Setup Write-Off"
                If Not ASCMAIN1.Logical_Lock("SPTDEMO1", "*") Then Exit Sub

            Case "Enter Payment"
                If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SPTDEMO1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SPTDCOM1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub

            Case "Update"

                If EntryMode = "W" And Not auto_writeoff Then

                    If txtWRITE_OFF_NOTE.Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Reference No"
                    End If

                    ' review the business rules below
                    If dst.Tables("SPTDCOMX").Select("AMT_TO_WOFF <> 0").Length = 0 Then
                        EMsg &= vbCr & "Nothing Written Off"
                    End If

                    If dst.Tables("SPTDCOMX").Select("AMT_TO_WOFF <> 0 and ((AMT_TO_WOFF < 0 AND AMT_COMM > 0) OR (AMT_TO_WOFF > 0 AND AMT_COMM < 0))").Length <> 0 Then
                        EMsg &= vbCr & "Some Write-Offs are not the same sign of the original accrual"
                    End If

                Else


                    If Absx1.txtFor("PYMT_REF_NO").Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Reference No"
                    End If

                    'If Absx1.dteFor("PYMT_REF_DATE").Value & "" = "" Then
                    '    EMsg &= vbCr & "You Must Specify a Reference Date"
                    'End If

                    Dim DT As Date = Absx1.dteFor("PYMT_REF_DATE").Value
                    If Absx1.dteFor("PYMT_REF_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Payment Date is Mandatory"
                    Else
                        If Now.Subtract(DT).TotalDays < 0 Then
                            EMsg &= vbCr & "Payment Date cannot be in the Future"
                        Else
                            If Now.Subtract(DT).TotalDays > 90 Then
                                EMsg &= vbCr & "Payment Date is more than 90 days in the past"
                            Else
                                If Now.Subtract(DT).TotalDays > 30 Then
                                    If MsgBox("Payment Date is more than 30 days prior to today", MsgBoxStyle.YesNo, "OK to Proceed?") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If

                    Dim CREDITED As Decimal = Val(numCREDITED.Value & "")
                    Dim AMT_TO_APPLY_total As Decimal = Val(dst.Tables("ARTOPENX").Compute("SUM(AMT_TO_APPLY)", "") & "")
                    If AMT_TO_APPLY_total = 0 And CREDITED <> 0 Then
                        EMsg &= vbCr & "You must apply the Payment amount to Chargebacks - why else would you be in this screen?"
                        ' AK SAYS SHE WILL NEVER USE THIS SCREEN WITHOUT A CHARGEBACK, AND SHE WILL NEVER WANT TO LEAVE A BALANCE OPEN ON THE CREDIT GENERATED - SHE WILL LEAVE THE BALANCE OPEN ON THE CHARGEBACKS INDIVIDUALLY

                    Else

                        If AMT_TO_APPLY_total > CREDITED Then
                            EMsg &= vbCr & "You cannot apply more than the total Credit Value (" & Format(CREDITED, "$#,##0.00") & ") to Chargebacks"
                        End If

                        Dim PYMT_REF_AMT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")
                        If PYMT_REF_AMT <> CREDITED Then
                            EMsg &= vbCr & "Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ") does not equal Total Credit distribution (" & Format(CREDITED, "$#,##0.00") & ")"
                        End If

                        If AMT_TO_APPLY_total <> PYMT_REF_AMT Then
                            EMsg &= vbCr & "Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ")" & vbCrLf & " does not equal the Total Chargebacks applied (" & Format(AMT_TO_APPLY_total, "$#,##0.00") & ")"
                            ' AK SAYS SHE WILL NEVER USE THIS SCREEN WITHOUT A CHARGEBACK, AND SHE WILL NEVER WANT TO LEAVE A BALANCE OPEN ON THE CREDIT GENERATED - SHE WILL LEAVE THE BALANCE OPEN ON THE CHARGEBACKS INDIVIDUALLY

                            ' note that the payments tab is based on ARTPYMT1/2/3 - so a demo payment without application to chargebacks won't even show up on that screen
                            ' so that screen probably needs to be re-architected to show records from SPTDPMT1

                            'If MsgBox("Chargebacks have been applied." & vbCrLf & vbCrLf & "However, the Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ")" & vbCrLf & " does not equal the Total Chargebacks applied (" & Format(AMT_TO_APPLY_total, "$#,##0.00") & ")." & vbCrLf _
                            '          & "An AR Item will remain on the Customer Account with a " & IIf(-1 * PYMT_REF_AMT + AMT_TO_APPLY_total > 0, "DR", "CR") & " Balance of " & Format(-1 * PYMT_REF_AMT + AMT_TO_APPLY_total, "$#,##0.00") & "." _
                            '                 & vbCrLf & vbCrLf & "OK to Continue with Payment?", _
                            '                 MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            '    Exit Sub
                            'End If
                        End If

                        For Each rowARTOPENX As DataRow In dst.Tables("ARTOPENX").Select("AMT_TO_APPLY<>0")
                            Dim AMT_TO_APPLY As Decimal = Val(rowARTOPENX.Item("AMT_TO_APPLY") & "")
                            Dim INV_BALANCE As Decimal = Val(rowARTOPENX.Item("INV_BALANCE") & "")
                            '  Dim INV_BALANCE_NEW As Decimal = Val(rowARTOPENX.Item("INV_BALANCE_NEW") & "")
                            If (INV_BALANCE >= 0 And (AMT_TO_APPLY < 0 Or AMT_TO_APPLY > INV_BALANCE)) _
                            Or (INV_BALANCE <= 0 And (AMT_TO_APPLY > 0 Or AMT_TO_APPLY < INV_BALANCE)) Then
                                EMsg &= vbCr & "Invalid Payment Amount for AR Item " & rowARTOPENX.Item("INV_TYPE") & "-" & rowARTOPENX.Item("INV_NUM")
                            End If
                        Next
                    End If

                    If EMsg = "" Then
                        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTDCOMC").Select("APPLY = '1'"), "OPS_YYYYPP").Select
                            Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP")
                            If dst.Tables("SPTDCOMC").Select("OPS_YYYYPP = '" & OPS_YYYYPP & "' and APPLY = '0'").Length <> 0 Then
                                If auto_writeoff Then
                                    ' don't stop to ask
                                Else
                                    If MsgBox("Some Accruals in " & OPS_YYYYPP & " have been queued for Payment leaving others behind" _
                                              & vbCrLf & vbCrLf & "OK to Continue?",
                                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If EMsg = "" Then
                        Dim msg As String = ""

                        msg &= vbCrLf & "- Generate a Credit Memo for " & Format(CREDITED, "$#,##0.00")
                        Dim DEMO_COMM_OFFSET As Decimal = Val(numDEMO_COMM_OFFSET.Value & "")
                        msg &= vbCrLf & "- Offset Demo Commission Accruals by " & Format(DEMO_COMM_OFFSET, "$#,##0.00")
                        Dim DEMO_COMM_ADD_EXP As Decimal = Val(numDEMO_COMM_ADD_EXP.Value & "")
                        msg &= vbCrLf & "- Record " & IIf(DEMO_COMM_ADD_EXP > 0, "additional", "a reduction to") & " Demo Commission Expenses of " & Format(DEMO_COMM_ADD_EXP, "$#,##0.00")
                        Dim DEDUCTED As Decimal = Val(numDEDUCTED.Value & "")
                        msg &= vbCrLf & "- Apply " & Format(DEDUCTED, "$#,##0.00") & " of the Credit Memo Generated against Open Demo Chargebacks"
                        msg &= vbCrLf & vbCrLf & "leaving the Credit Memo Generated with an Open Balance of " & Format(CREDITED - DEDUCTED, "$#,##0.00")

                        Dim BALANCE As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(BALANCE)", "APPLY='1' AND BALANCE<>0") & "")
                        If BALANCE <> 0 Then
                            msg &= vbCrLf & vbCrLf & " and leave payment variances of " & Format(BALANCE, "$#,##0.00") & " as open accruals"
                        End If
                        If auto_writeoff Then
                            ' don't stop to ask
                        Else
                            If MsgBox("You are about to:" & vbCrLf & msg & vbCrLf & vbCrLf & "OK to Proceed?",
                                MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Reverse"

                If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SPTDEMO1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SPTDCOM1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub

                If MsgBox("*** This Action is NOT Reversible ***" & vbCrLf & vbCrLf & "Are you Sure that you want to Reverse this Payment?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
        End Select

        If EMsg <> "" Then
            auto_writeoff = False
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Enter Payment"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "View Payment"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Setup Write-Off"
                EntryMode = "W"
                ' Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "W" Then
                    Update_Record_Write_Off
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Reverse"
                Reverse_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Setup Write-Off").Settings.Enabled = not_iScreenMode

                    .Items("Enter Payment").Settings.Enabled = IIf(ScreenMode And Not (EntryMode = "N"), _
                                                         DefaultableBoolean.True, _
                                                         DefaultableBoolean.False)

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Reverse").Visible = Not (EntryMode = "N") And Not (EntryMode = "W") _
                        AndAlso (ScreenMode And EntryMode <> "L") AndAlso (rowSPTDPMT1.Item("REVERSED_BY_PYMT_NO") & "" = "" And rowSPTDPMT1.Item("REVERSED_PYMT_NO") & "" = "")

                    .Items("Enter Payment").Visible = ScreenMode And (EntryMode = "L")
                    .Items("Done").Visible = Not (EntryMode = "N") And Not (EntryMode = "W")
                    .Items("Update").Visible = ScreenMode And ((EntryMode = "N") Or (EntryMode = "W"))
                    .Items("Cancel").Visible = ScreenMode And ((EntryMode = "N") Or (EntryMode = "W"))
                End With

                .Groups("Totals").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")
                .Groups("Customer Document").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(grpCustomerDocument, (EntryMode <> "N"))

        If EntryMode = "W" Then
            tab0.SelectedTab = tab0.Tabs("Open Accruals")
        End If

        grdSPTDCOMX.DisplayLayout.Bands(0).Columns("WRITE_OFF").Hidden = Not (EntryMode = "W")
        grdSPTDCOMX.DisplayLayout.Bands(0).Columns("AMT_TO_WOFF").Hidden = Not (EntryMode = "W")
        grdSPTDCOMX.DisplayLayout.Bands(0).Columns("AMT_COMM_NEW").Hidden = Not (EntryMode = "W")

        For itab As Integer = 0 To tab0.Tabs.Count - 1
            If tab0.Tabs(itab).Key = "Open Accruals" Then
            Else
                tab0.Tabs(itab).Enabled = Not (EntryMode = "W")
            End If
        Next

        If EntryMode = "W" Then
            'spl.Panel1Collapsed = True
            'grdSPTDCOMX.Text = "Open Demo Accruals - Setup Write-Off"
            lblCUST_CODE.Visible = False
            lblCUST_NAME.Visible = False
            Absx1.txtFor("CUST_CODE").Visible = False
            Absx1.txtFor("CUST_NAME").Visible = False
            lblWriteOff.Visible = True

            lblWRITE_OFF_NOTE.Visible = True
            txtWRITE_OFF_NOTE.Visible = True

            Set_Read_Only_for_ctl(txtWRITE_OFF_NOTE, False)
            btnWriteOff.Visible = False

            Dim DVW As DataView = DirectCast(grdSPTDCOMX.DataSource, DataTable).DefaultView
            DVW.RowFilter = "AMT_COMM <> 0"
        Else
            grpStores.Visible = Not tf
            'spl.Panel1Collapsed = False
            'grdSPTDCOMX.Text = "Open Demo Accruals"
            lblCUST_CODE.Visible = True
            lblCUST_NAME.Visible = True
            Absx1.txtFor("CUST_CODE").Visible = True
            Absx1.txtFor("CUST_NAME").Visible = True
            lblWriteOff.Visible = False

            grpARTCUST1.Visible = tf
            Set_Read_Only(grpARTCUST1, True)

            lblWRITE_OFF_NOTE.Visible = False
            txtWRITE_OFF_NOTE.Visible = False

        End If



        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdSPTDCOMC, grdSPTDCOML, grdSPTDCOMP, grdSPTDCOMS, grdSPTDCOMX, grdARTOPENX, grdSPTDPMTX}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            If grd.Name = "grdSPTDCOMC" Or grd.Name = "grdARTOPENX" Then
                If EntryMode = "N" Then
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            End If
            If grd.Name = "grdSPTDCOMX" Then
                If EntryMode = "W" Then
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            End If
        Next


        With grdARTOPENX
            .Parent = IIf(ScreenMode, _
                         tabDetails.Tabs("Open Chargebacks").TabPage, _
                         tab0.Tabs("Open Chargebacks").TabPage)
            With .DisplayLayout.Bands(0)
                .Columns("PAY").Hidden = (Not ScreenMode Or EntryMode <> "N")
                .Columns("AMT_TO_APPLY").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
                .Columns("INV_BALANCE_NEW").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
                .Columns("CUST_CODE").Hidden = ScreenMode
            End With
        End With

        With grdSPTDCOMC.DisplayLayout.Bands(0)
            .Columns("AMT_SOLD_CLAIMED").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
            .Columns("AMT_COMM_PAID").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
            .Columns("APPLY").Hidden = (Not ScreenMode Or EntryMode <> "N")
            .Columns("AMT_COMM_ADD").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
            .Columns("AMT_COMM_RED").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
            .Columns("LEAVE_OPEN").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
            .Columns("BALANCE").Hidden = (Not ScreenMode Or (EntryMode <> "N" And EntryMode <> "V"))
        End With

        lblPYMT_NO.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")
        txtPYMT_NO.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")

        If ScreenMode Then
            If EntryMode = "V" Then
                grdARTOPENX.Text = "Chargebacks Applied to this Demo Payment"
                grdSPTDCOMC.Text = "Demo Commission Accruals used in this Demo Payment"
                tabDetails.Tabs("Open Chargebacks").Text = "Chargebacks Applied"
                tabDetails.Tabs("Paid Previously").Visible = False
            Else
                grdARTOPENX.Text = "Customer Deductions Charged Back"
                grdSPTDCOMC.Text = "Open Demo Commission Accruals to Apply"
                tabDetails.Tabs("Open Chargebacks").Text = "Open Chargebacks"
                tabDetails.Tabs("Paid Previously").Visible = True
            End If

            lblStatus.Visible = False
            If EntryMode = "V" Then
                If rowSPTDPMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Then
                    lblStatus.Text = "Reversed by " & rowSPTDPMT1.Item("REVERSED_BY_PYMT_NO")
                ElseIf rowSPTDPMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                    lblStatus.Text = "Reversing " & rowSPTDPMT1.Item("REVERSED_PYMT_NO")
                Else
                    lblStatus.Text = "Issued " & Format(rowSPTDPMT1.Item("INIT_DATE"), "MM/dd/yy") & " by " & rowSPTDPMT1.Item("INIT_OPER")
                End If
                lblStatus.Visible = True
            ElseIf EntryMode = "N" Then
                lblStatus.Text = "In Process"
                lblStatus.Visible = True
            End If

            lblPYMT_CTL_NO.Visible = (EntryMode <> "N")

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTDCOMC", "SPTDCOMP", "ARTOPENX", "SPTDPMT1", "SPTDPMT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_Open_Accruals()
        Absx1.txtFor("CUST_CODE").Text = ""
        'Absx1.txtFor("PYMT_REF_NO").Text = ""
        'Absx1.dteFor("PYMT_REF_DATE").Value = DBNull.Value
        'Absx1.numFor("PYMT_REF_AMT").Value = 0

        ACC_CTL_NOs.Clear()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = HFs("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)


        Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)


        If EntryMode = "N" Then
            PYMT_NO = ASCMAIN1.Next_Control_No("SPTDPMT1.PYMT_NO")
            rowSPTDPMT1 = dst.Tables("SPTDPMT1").NewRow
            rowSPTDPMT1.Item("PYMT_NO") = PYMT_NO
            rowSPTDPMT1.Item("CUST_CODE") = CUST_CODE
            rowSPTDPMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSPTDPMT1.Item("INIT_DATE") = DATETIME_STAMP
            rowSPTDPMT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("SPTDPMT1").Rows.Add(rowSPTDPMT1)
        ElseIf EntryMode = "V" Then
            rowSPTDPMT1 = Fill_Record("SPTDPMT1", PYMT_NO)
        End If

        EnforceConstraints(False)

        If EntryMode = "N" Or EntryMode = "L" Then
            Fill_Records("SPTDCOMC", CUST_CODE)
            Fill_Records("SPTDCOMP", CUST_CODE)
        ElseIf EntryMode = "V" Then
            ASCMAIN1.sql = "Select * from SPTDCOMC where PYMT_NO = '" & PYMT_NO & "'"
            Fill_Records("SPTDCOMC", "", True, ASCMAIN1.sql)

            If rowSPTDPMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                For Each row As DataRow In dst.Tables("SPTDCOMC").Select("")
                    ' PROBABLY SHOULD HAVE ACC_CTL_NO_OPEN REFERENCE IN SPTDCOMC TO INDICATE THAT BALANCE WAS LEFT OPEN
                    Dim AMT_COMM_ADJ As Decimal = Val(row.Item("AMT_COMM_ADJ") & "")
                    'Dim R As Int32 = dst.Tables("SPTDCOMC").Rows.Count
                    'Dim ACC_CTL_NO As String = row.Item("ACC_CTL_NO")
                    'If ASCDATA1.GetDataRow("Select * from SPTDCOMC where ACC_CTL_NO_ORIG = '" & ACC_CTL_NO & "'") IsNot Nothing Then
                    '    row.Item("LEAVE_OPEN") = "1"
                    'End If
                    If AMT_COMM_ADJ = 0 Then
                        row.Item("LEAVE_OPEN") = "1"
                    End If
                Next
            End If

            Fill_Records("SPTDCOMP", "", True, ASCMAIN1.sql)
        End If

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_DEMO_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        If EntryMode = "N" Or EntryMode = "L" Then
            Fill_Records("ARTOPENX", New String() {REASON_CODE, CUST_BILL_TO_CUST})
        ElseIf EntryMode = "V" Then
            PYMT_BATCH_NO = rowSPTDPMT1.Item("PYMT_BATCH_NO") & ""

            Dim CUST_CODE_PYMT As String = CUST_BILL_TO_CUST
            ASCMAIN1.sql = "Select CUST_CODE from ARTPYMT2 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            Dim rowC() As DataRow = ASCDATA1.GetDataTable.Select("")
            If rowC.Length = 1 Then
                CUST_CODE_PYMT = rowC(0).Item(0)
            End If

            ASCMAIN1.sql = sqlARTOPENX & " from ARTOPEN1 where CUST_CODE = '" & CUST_CODE_PYMT & "' and (INV_TYPE,INV_NUM) in (Select INV_TYPE,INV_NUM from ARTPYMT3 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "' and PYMT_BATCH_ILNO <> 0)"
            'ASCMAIN1.sql = sqlARTOPENX & " from ARTOPEN1 where (INV_TYPE,INV_NUM) in (Select INV_TYPE,INV_NUM from ARTPYMT3 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "' and PYMT_BATCH_ILNO <> 0)"
            Fill_Records("ARTOPENX", "", True, ASCMAIN1.sql)
            Fill_Records("ARTOPENX", "", False, Replace(ASCMAIN1.sql, "from ARTOPEN1", "from ARTOPENX ARTOPEN1"))

            ASCMAIN1.sql = "Select * from ARTPYMT3 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("PYMT_BATCH_ILNO <> 0")
                Dim INV_TYPE As String = row.Item("INV_TYPE")
                Dim INV_NUM As String = row.Item("INV_NUM")
                'Dim rowARTOPENX As DataRow = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_BILL_TO_CUST, INV_TYPE, INV_NUM})
                Dim rowARTOPENX As DataRow = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_CODE_PYMT, INV_TYPE, INV_NUM})
                'If rowARTOPENX Is Nothing Then
                '    ASCMAIN1.sql = sqlARTOPENX & $" from ARTOPEN1 where CUST_CODE = '{CUST_BILL_TO_CUST}' and INV_TYPE = '{INV_TYPE}' and INV_NUM = '{INV_NUM}'"
                '    Fill_Records("ARTOPENX", "", False, ASCMAIN1.sql)
                '    rowARTOPENX = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_BILL_TO_CUST, INV_TYPE, INV_NUM})
                'End If

                rowARTOPENX.Item("INV_BALANCE") = row.Item("INV_BALANCE")
                    rowARTOPENX.Item("AMT_TO_APPLY") = row.Item("INV_PMT")
            Next

            For Each row As DataRow In dst.Tables("SPTDCOMC").Select("")
                row.Item("APPLY") = "1"
            Next

        End If

        EnforceConstraints(True)

        With dst.Tables("SPTDCOMC")
            If rowSPTDPMT1 IsNot Nothing AndAlso EntryMode <> "L" AndAlso rowSPTDPMT1.Item("REVERSED_PYMT_NO") & "" = "" Then
                ' NORMAL
                .Columns("AMT_COMM_ADD_CALC").Expression = "IIF(APPLY='0' OR ISNULL(AMT_COMM,0)=ISNULL(AMT_COMM_PAID,0) OR ISNULL(AMT_COMM,0)>ISNULL(AMT_COMM_PAID,0),NULL,ISNULL(AMT_COMM_PAID,0)-ISNULL(AMT_COMM,0))"
                .Columns("AMT_COMM_RED_CALC").Expression = "IIF(APPLY='0' OR ISNULL(AMT_COMM,0)=ISNULL(AMT_COMM_PAID,0) OR ISNULL(AMT_COMM,0)<ISNULL(AMT_COMM_PAID,0),NULL,ISNULL(AMT_COMM,0)-ISNULL(AMT_COMM_PAID,0))"
                .Columns("BALANCE").Expression = "IIF(LEAVE_OPEN='0',NULL,ISNULL(AMT_COMM_RED_CALC,0)-ISNULL(AMT_COMM_ADD_CALC,0))"
            Else
                ' REVERSE PYMT
                .Columns("AMT_COMM_ADD_CALC").Expression = "IIF(ISNULL(AMT_COMM_ADJ,0)>0,AMT_COMM_ADJ,0)"
                .Columns("AMT_COMM_RED_CALC").Expression = "IIF(ISNULL(AMT_COMM_ADJ,0)<0,-1 * AMT_COMM_ADJ,0)"
                .Columns("BALANCE").Expression = "-1 * ISNULL(AMT_COMM_OFFSET,0)"
            End If

        End With


        Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record_Write_Off()

        ' Stop
        Dim WRITE_OFF_NOTE As String = txtWRITE_OFF_NOTE.Text


        Dim tbl As DataTable = dst.Tables("SPTDCOMX").Select("WRITE_OFF = '1' AND AMT_TO_WOFF <> 0").CopyToDataTable()
        tbl.PrimaryKey = New DataColumn() {tbl.Columns("ACC_CTL_NO")}
        Dim AMT_TO_WOFF_total As Decimal = Val(tbl.Compute("SUM(AMT_TO_WOFF)", "") & "")
        If MsgBox("OK to Write Off " & Format(AMT_TO_WOFF_total, "$#,##0.00") & "?", MsgBoxStyle.OkCancel, "Verification") <> MsgBoxResult.Ok Then Exit Sub

        auto_writeoff = True
        For Each row As DataRow In ASCDATA1.SelectDistinct(tbl, "CUST_CODE").Select("", "CUST_CODE")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim RECS As Int64 = Val(tbl.Select($"CUST_CODE = '{CUST_CODE}'").Length)
            Dim AMT_COMM As Decimal = Val(tbl.Compute("SUM(AMT_COMM)", $"CUST_CODE = '{CUST_CODE}'") & "")
            Dim AMT_TO_WOFF As Decimal = Val(tbl.Compute("SUM(AMT_TO_WOFF)", $"CUST_CODE = '{CUST_CODE}'") & "")

            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Click_Command("View")
            Click_Command("Enter Payment")

            rowSPTDPMT1.Item("PYMT_REF_DATE") = Now.Date
            rowSPTDPMT1.Item("PYMT_REF_NO") = WRITE_OFF_NOTE

            For Each grow As UltraWinGrid.UltraGridRow In grdSPTDCOMC.Rows
                Dim ACC_CTL_NO As String = grow.Cells("ACC_CTL_NO").Value & ""
                Dim row2 As DataRow = tbl.Rows.Find(New String() {ACC_CTL_NO})
                If row2 IsNot Nothing Then
                    grow.Cells("APPLY").Value = "1"
                    grow.Cells("AMT_COMM_PAID").Value = 0
                    grow.Update()
                End If
            Next

            If Val(Absx1.numFor("DEMO_COMM_OFFSET").Value & "") <> AMT_COMM Then
                Debug.Print(CUST_CODE & ":" & rowSPTDPMT1.Item("PYMT_NO") & ":" & Val(rowSPTDPMT1.Item("DEMO_COMM_OFFSET") & "") & ":" & CStr(AMT_COMM))
                ' Stop
                'Throw New Exception("If Val(Absx1.numFor('DEMO_COMM_OFFSET').Value & '') <> AMT_COMM Then")
                'Absx1.numFor("DEMO_COMM_OFFSET").Value = AMT_COMM
                rowSPTDPMT1.Item("DEMO_COMM_OFFSET") = AMT_COMM
                rowSPTDPMT1.Item("DEMO_COMM_ADD_EXP") = -1 * AMT_COMM
                'row.AcceptChanges()
                'Stop
                'Dim I As Integer = dst.Tables("SPTDPMT1").Rows.Count
                'Stop
            End If


            Click_Command("Update")
            If Not auto_writeoff Then
                Exit For
            End If
        Next

        auto_writeoff = False
        MsgBox("Write Off is Complete", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        'If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        'Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        dst.Tables("SPTDCOMB").Rows.Clear()
        dst.Tables("SPTDCOMC").AcceptChanges()

        Dim OPS_YYYYPP As String = ""

        Dim DEMO_COMM_OFFSET As Decimal = 0
        Dim DEMO_COMM_CREDITED As Decimal = 0

        For Each rowSPTDCOMC As DataRow In dst.Tables("SPTDCOMC").Select("LEAVE_OPEN = '1' AND BALANCE <> 0")
            Dim ACC_CTL_NO_orig As String = rowSPTDCOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO As String = Replicate_Accrual(ACC_CTL_NO_orig, True)
        Next

        For Each rowSPTDCOMC As DataRow In dst.Tables("SPTDCOMC").Select("APPLY = '1'", "CUST_CODE,OPS_YYYYPP")

            Dim ACC_CTL_NO As String = rowSPTDCOMC.Item("ACC_CTL_NO")
            If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                rowSPTDCOMC.SetAdded()
            End If

            rowSPTDCOMC.Item("OPS_YYYYPP_PAID") = ASCMAIN1.CYP
            rowSPTDCOMC.Item("PYMT_NO") = PYMT_NO

            If OPS_YYYYPP <> rowSPTDCOMC.Item("OPS_YYYYPP") Then
                OPS_YYYYPP = rowSPTDCOMC.Item("OPS_YYYYPP")
                Fill_Records("SPTDCOMB", New String() {CUST_CODE, OPS_YYYYPP}, False)
            End If

            Dim HC_CODE As String = rowSPTDCOMC.Item("HC_CODE")
            Dim AMT_COMM_total As Decimal = Val(rowSPTDCOMC.Item("AMT_COMM") & "")
            Dim AMT_COMM_PAID_total As Decimal = Val(rowSPTDCOMC.Item("AMT_COMM_PAID") & "")

            DEMO_COMM_CREDITED += AMT_COMM_PAID_total

            If rowSPTDCOMC.Item("LEAVE_OPEN") & "" = "1" Then
                DEMO_COMM_OFFSET += AMT_COMM_PAID_total
                rowSPTDCOMC.Item("AMT_COMM_ADJ") = 0
                rowSPTDCOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_PAID_total
            Else
                DEMO_COMM_OFFSET += AMT_COMM_total
                rowSPTDCOMC.Item("AMT_COMM_ADJ") = AMT_COMM_PAID_total - AMT_COMM_total
                rowSPTDCOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_total
            End If

            Dim AMT_COMM_PAID_spread As Decimal = 0

            Dim sqlw As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and CUST_CODE = '" & CUST_CODE & "' and HC_CODE = '" & HC_CODE & "'"
            Dim AMT_COMM_all As Decimal = Val(dst.Tables("SPTDCOMB").Compute("SUM(AMT_COMM)", sqlw) & "")

            Dim rows() As DataRow = dst.Tables("SPTDCOMB").Select(sqlw, "AMT_COMM")
            Dim r As Integer = 0

            Dim COLLECTION_CODEs As New Dictionary(Of String, Decimal)

            For Each rowSPTDCOMB As DataRow In rows

                Dim AMT_COMM As Decimal = Val(rowSPTDCOMB.Item("AMT_COMM") & "")
                Dim AMT_COMM_PAID As Decimal = Val(rowSPTDCOMB.Item("AMT_COMM_PAID") & "")

                Dim F As Decimal = 1
                If AMT_COMM_all <> 0 Then F = AMT_COMM / AMT_COMM_all

                Dim AMT_COMM_PAID_now As Decimal = System.Math.Round(AMT_COMM_PAID_total * F, 2)
                AMT_COMM_PAID_spread += AMT_COMM_PAID_now
                AMT_COMM_PAID += AMT_COMM_PAID_now

                r += 1
                If r = rows.Length And AMT_COMM_PAID_total <> AMT_COMM_PAID_spread Then
                    AMT_COMM_PAID += AMT_COMM_PAID_total - AMT_COMM_PAID_spread
                End If

                rowSPTDCOMB.Item("AMT_COMM_PAID") = AMT_COMM_PAID

                If AMT_COMM_PAID <> AMT_COMM Then
                    Dim COLLECTION_CODE As String = rowSPTDCOMB.Item("COLLECTION_CODE")
                    If Not COLLECTION_CODEs.ContainsKey(COLLECTION_CODE) Then
                        COLLECTION_CODEs.Add(COLLECTION_CODE, 0)
                    End If
                    If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                        COLLECTION_CODEs(COLLECTION_CODE) += AMT_COMM_PAID_now
                    Else
                        COLLECTION_CODEs(COLLECTION_CODE) += AMT_COMM_PAID - AMT_COMM
                    End If
                End If
            Next

            ' Record impact to Commission Adjustment (expense)

            If COLLECTION_CODEs.Count <> 0 Then
                For Each COLLECTION_CODE As String In COLLECTION_CODEs.Keys
                    If COLLECTION_CODEs(COLLECTION_CODE) <> 0 Then
                        Dim rowSPTDPMT2 As DataRow = dst.Tables("SPTDPMT2").NewRow
                        rowSPTDPMT2.Item("PYMT_NO") = PYMT_NO
                        rowSPTDPMT2.Item("ACC_CTL_NO") = rowSPTDCOMC.Item("ACC_CTL_NO")
                        rowSPTDPMT2.Item("COLLECTION_CODE") = COLLECTION_CODE
                        If rowSPTDCOMC.Item("LEAVE_OPEN") & "" = "1" Then
                            rowSPTDPMT2.Item("AMT_COMM_ADJ") = 0
                        Else
                            rowSPTDPMT2.Item("AMT_COMM_ADJ") = COLLECTION_CODEs(COLLECTION_CODE)
                        End If
                        dst.Tables("SPTDPMT2").Rows.Add(rowSPTDPMT2)
                    End If
                Next
            End If
        Next

        Dim PYMT_REF_DATE As Date = Absx1.dteFor("PYMT_REF_DATE").Value
        Dim PYMT_REF_NO As String = Absx1.txtFor("PYMT_REF_NO").Text

        'Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_DEMO_TYPE_CODE")
        'Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)

        'Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
        'Dim TERM_CODE As String = rowSOTTYPE1.Item("TERM_CODE")
        'Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        Dim PYMT_CTL_NO As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        rowSPTDPMT1.Item("PYMT_CTL_NO") = PYMT_CTL_NO

        rowSPTDPMT1.Item("DEMO_COMM_OFFSET") = DEMO_COMM_OFFSET
        rowSPTDPMT1.Item("DEMO_COMM_CREDITED") = DEMO_COMM_CREDITED


        Dim DEMO_COMM_APPLIED As Decimal = Val(Absx1.numFor("DEDUCTED").Value & "")
        Dim DEMO_COMM_ADD_EXP As Decimal = Val(Absx1.numFor("DEMO_COMM_ADD_EXP").Value & "")

        rowSPTDPMT1.Item("DEMO_COMM_APPLIED") = DEMO_COMM_APPLIED

        If auto_writeoff Then
            rowSPTDPMT1.Item("DEMO_COMM_ADD_EXP") = -1 * Val(rowSPTDPMT1.Item("DEMO_COMM_OFFSET") & "")
        Else
            rowSPTDPMT1.Item("DEMO_COMM_ADD_EXP") = DEMO_COMM_ADD_EXP
        End If

        rowSPTDPMT1.Item("PYMT_DATE") = PYMT_REF_DATE



        Dim INV_TOTAL_AMOUNT As Decimal = Val(Absx1.numFor("CREDITED").Value & "")
        Dim DEDUCTED As Decimal = Val(Absx1.numFor("DEDUCTED").Value & "")
        Dim INV_DATE As Date = Now.Date

        rowSPTDPMT1.Item("PYMT_BATCH_NO") = Setup_ARTPYMTx(PYMT_CTL_NO, PYMT_REF_DATE, PYMT_REF_NO, INV_DATE, INV_TOTAL_AMOUNT, DEDUCTED)

        ' ******************* Start of ARTPYMTx 

        'Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        'With rowARTOPEN1
        '    .Item("CUST_CODE") = CUST_BILL_TO_CUST
        '    .Item("INV_TYPE") = "C"
        '    .Item("INV_NUM") = PYMT_CTL_NO
        '    .Item("INV_DATE") = INV_DATE
        '    .Item("POST_CODE") = POST_CODE
        '    .Item("TERM_CODE") = TERM_CODE
        '    .Item("INV_DUE_DATE") = INV_DATE
        '    .Item("INV_CUST_PO") = PYMT_REF_NO
        '    .Item("INV_MISC_CHG") = -1 * INV_TOTAL_AMOUNT
        '    .Item("INV_TOTAL_AMOUNT") = -1 * INV_TOTAL_AMOUNT
        '    .Item("INV_BALANCE") = -1 * INV_TOTAL_AMOUNT
        '    .Item("CUST_CODE_SO") = CUST_CODE
        '    .Item("REASON_CODE") = REASON_CODE
        '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '    .Item("INIT_DATE") = DATETIME_STAMP
        '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP

        '    .Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE

        '    .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
        '    .Item("SALES_DIVISION_CODE") = ""

        '    .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        '    .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        '    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        '    .Item("INV_SALES_CURR") = .Item("INV_SALES")
        '    .Item("INV_DISC_CURR") = .Item("INV_DISC")
        '    .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
        '    .Item("INV_STAX_CURR") = .Item("INV_STAX")
        '    .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")
        '    .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
        '    .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")

        '    .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        '    .Item("CURR_EXCH_RATE") = 1
        'End With

        'dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        'If dst.Tables("ARTOPENX").Select("AMT_TO_APPLY <> 0").Length <> 0 Then

        '    Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        '    rowSPTDPMT1.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO

        '    Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        '    With rowARTPYMT1
        '        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
        '        .Item("PYMT_BATCH_DATE") = INV_DATE
        '        .Item("OPS_YYYYPP") = ASCMAIN1.CYP
        '        .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '        .Item("INIT_DATE") = DATETIME_STAMP
        '        .Item("LAST_OPER") = ASCMAIN1.USER_ID
        '        .Item("LAST_DATE") = DATETIME_STAMP
        '        .Item("STATUS") = "1"
        '        .Item("PYMT_APPL_ONLY") = "1"
        '        .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        '        .Item("CURR_EXCH_RATE") = 1
        '        .Item("PYMT_SOURCE") = "D"
        '    End With
        '    dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        '    Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        '    With rowARTPYMT2
        '        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
        '        .Item("PYMT_BATCH_LNO") = 1
        '        .Item("CUST_CODE") = CUST_BILL_TO_CUST
        '        .Item("CUST_NAME") = rowARTCUST1_BT.Item("CUST_NAME")
        '        .Item("CUST_PYMT_REF_NO") = PYMT_REF_NO
        '        .Item("CUST_PYMT_REF_DATE") = PYMT_REF_DATE
        '        .Item("CUST_PYMT_AMT") = 0
        '        .Item("PYMT_STATUS") = "2"
        '        .Item("CUST_PYMT_AMT_CURR") = 0
        '        .Item("PYMT_NOTE") = "Demo Payment"
        '        .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '        .Item("INIT_DATE") = DATETIME_STAMP
        '        .Item("LAST_OPER") = ASCMAIN1.USER_ID
        '        .Item("LAST_DATE") = DATETIME_STAMP
        '    End With
        '    dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

        '    dst.Tables("ARTPYMT3").Rows.Clear()

        '    ' Record Credit in ARTPYMT3
        '    Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, PYMT_BATCH_NO, 0, -1 * DEDUCTED, PYMT_REF_NO, PYMT_REF_DATE)

        '    ' Record Payments against Chargebacks
        '    Dim PYMT_BATCH_ILNO As Integer = 0
        '    For Each rowARTOPENX As DataRow In dst.Tables("ARTOPENX").Select("AMT_TO_APPLY <> 0", "INV_NUM")
        '        Dim CUST_CODE As String = rowARTOPENX.Item("CUST_CODE")
        '        Dim INV_TYPE As String = rowARTOPENX.Item("INV_TYPE")
        '        Dim INV_NUM As String = rowARTOPENX.Item("INV_NUM")
        '        Dim INV_PMT As Decimal = Val(rowARTOPENX.Item("AMT_TO_APPLY") & "")
        '        rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {CUST_CODE, INV_TYPE, INV_NUM}, False, False)
        '        PYMT_BATCH_ILNO += 1
        '        Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, PYMT_BATCH_NO, PYMT_BATCH_ILNO, INV_PMT, PYMT_REF_NO, PYMT_REF_DATE)
        '    Next
        'End If


        ' ******************* End of ARTPYMTx 

        ASCMAIN1.Record_Event("SPTDPMT1", PYMT_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "P", "Payment Entered", "")

        Update_Record_TDA("SPTDCOMC")
        Update_Record_TDA("SPTDCOMB")

        Update_Record_TDA("SPTDPMT1")
        Update_Record_TDA("SPTDPMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTOPEN1")

        If auto_writeoff Then
            CommitTrans("")
        Else
            CommitTrans("Update Complete")
        End If

    End Sub

    Function Write_ARTPYMT3_from_ARTOPEN1( _
                                         row As DataRow, _
                                         PYMT_BATCH_NO As String, _
                                         PYMT_BATCH_ILNO As Integer, _
                                         INV_PMT As Decimal, _
                                         PYMT_REF_NO As String, _
                                         PYMT_REF_DATE As Date) As DataRow

        Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
        With rowARTPYMT3
            '"INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF", 
            For Each COLUMN_NAME As String In New String() _
                {"INV_TYPE", "INV_NUM", "REASON_CODE", "INV_DATE", "INV_DUE_DATE", "CUST_CODE_SO", _
                 "CUST_STORE_NO", "INV_CUST_PO", "INV_BALANCE", _
                  "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ORDR_TYPE_CODE", "CUST_CODE"}
                .Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
            Next
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_LNO") = 1
            .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
            .Item("INV_PMT") = INV_PMT
            .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE") & "") - INV_PMT

            .Item("INV_BALANCE_CURR") = Val(.Item("INV_BALANCE") & "")
            .Item("INV_PMT_CURR") = .Item("INV_PMT")
            .Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
            .Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
            .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW")
        End With

        With row
            .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT
            .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - INV_PMT
            .Item("INV_LAST_PMT") = DATETIME_STAMP.Date
            .Item("INV_LAST_PMT_REF") = PYMT_REF_NO
            .Item("INV_LAST_PMT_REF_DT") = PYMT_REF_DATE

            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP

            .Item("INV_PMT_CURR") = .Item("INV_PMT")
            .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
        End With

        dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)

        Return rowARTPYMT3
    End Function

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub


    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View"
                Absx1.txtFor("PYMT_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SPTDPMT1"
            E.COLUMN_NAME = "PYMT_NO"
            E.CODE_VALUE = Absx1.txtFor("PYMT_NO").Text
            E.DESC_VALUE = "Demo Comm Payment"
            E.ATTACHMENT_NOTES = ""
            'If rowAPTINVH1.Item("INV_STATUS") & "" <> "O" And rowAPTINVH1.Item("INV_STATUS") & "" <> "H" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "SPTDPMT1"
        E.TABLE_KEY_CAPTION = "Demo Comm Pymt No"
        If ScreenMode Then
            E.enabled = (EntryMode = "N")
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PYMT_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("CUST_CODE").Text & " " & Absx1.txtFor("CUST_NAME").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "V")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function
#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTDCOMX, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Write-Off All Visible")
        Load_Popup_Menu(grdSPTDCOMC, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Apply Selected", "Replicate")
        Load_Popup_Menu(grdSPTDCOMP, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Replicate")
        Load_Popup_Menu(grdARTOPENX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSPTDPMTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
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
            Case "grdSPTDCOMC"
                tlb_pop.Tools("Apply Selected").SharedProps.Visible = (EntryMode = "N")
                tlb_pop.Tools("Replicate").SharedProps.Visible = (EntryMode = "N")

            Case "grdSPTDCOMP"
                tlb_pop.Tools("Replicate").SharedProps.Visible = (EntryMode = "N")

            Case "grdSPTDCOMX"
                tlb_pop.Tools("Write-Off All Visible").SharedProps.Visible = (EntryMode = "W")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSPTDCOMX"
                    'If grdSPTDCOMX.Tag = "" Then
                    '    e.Cancel = True
                    'End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Apply Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("AMT_COMM_PAID").Value = grow.Cells("AMT_COMM").Value
                    grow.Update()
                Next
                grd.Selected.Rows.Clear()

            Case "Write-Off All Visible"
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTDCOMX.Rows
                    If grow.IsFilteredOut OrElse Val(grow.Cells("AMT_COMM").Value & "") = 0 Then
                        ' DO NOTHING
                    Else
                        Dim ACC_CTL_NO As String = grow.Cells("ACC_CTL_NO").Value
                        Dim rowSPTDCOMX As DataRow = dst.Tables("SPTDCOMX").Rows.Find(ACC_CTL_NO)
                        rowSPTDCOMX.Item("WRITE_OFF") = "1"
                        rowSPTDCOMX.Item("AMT_TO_WOFF") = grow.Cells("AMT_COMM").Value
                        'grow.Cells("WRITE_OFF").Value = "1"
                        'grow.Cells("AMT_TO_WOFF").Value = grow.Cells("AMT_COMM").Value
                        'grow.Update()
                    End If
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Replicate"
                Dim ACC_CTL_NO_orig As String = grd.ActiveRow.Cells("ACC_CTL_NO").Value
                Dim ACC_CTL_NO As String = Replicate_Accrual(ACC_CTL_NO_orig)

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow Then
                        If grow.Cells("ACC_CTL_NO").Value & "" = ACC_CTL_NO Then
                            grow.Activate()
                            Exit For
                        End If
                    End If
                Next
        End Select
    End Sub

#End Region

    Function Replicate_Accrual(ACC_CTL_NO_orig As String, Optional leave_open As Boolean = False) As String
        Dim rowSPTDCOMC_orig As DataRow
        If leave_open Then
            rowSPTDCOMC_orig = dst.Tables("SPTDCOMC").Rows.Find(ACC_CTL_NO_orig)
        Else
            rowSPTDCOMC_orig = dst.Tables("SPTDCOMP").Rows.Find(ACC_CTL_NO_orig)
        End If
        Dim ACC_CTL_NO As String = ASCMAIN1.Next_Control_No("SPTDCOMC.ACC_CTL_NO")
        ACC_CTL_NOs.Add(ACC_CTL_NO)

        Dim rowSPTDCOMC As DataRow = dst.Tables("SPTDCOMC").NewRow
        rowSPTDCOMC.ItemArray = rowSPTDCOMC_orig.ItemArray
        With rowSPTDCOMC
            .Item("ACC_CTL_NO") = ACC_CTL_NO
            If leave_open Then
                .Item("AMT_COMM") = rowSPTDCOMC_orig.Item("BALANCE")
                .Item("APPLY") = "0"
            Else
                .Item("AMT_COMM") = 0
                .Item("APPLY") = "1"
            End If

            .Item("AMT_COMM_OFFSET") = 0
            .Item("AMT_SOLD_CLAIMED") = 0
            .Item("AMT_COMM_PAID") = 0

            .Item("ACC_CTL_NO_ORIG") = ACC_CTL_NO_orig
            .Item("PYMT_NO_ORIG") = PYMT_NO

            .Item("OPS_YYYYPP_PAID") = DBNull.Value
        End With
        dst.Tables("SPTDCOMC").Rows.Add(rowSPTDCOMC)

        Return ACC_CTL_NO
    End Function
#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdSPTDCOMX"

    Private Sub grdSPTDCOMX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTDCOMX.AfterCellUpdate
        'If e.Cell.Column.Key = "CUST_DC_IND" Then
        '    'grdSPTDCOMX.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        '    grdSPTDCOMX.UpdateData()
        'End If

        'If e.Cell.Value & "" <> "" Then

        'End If

    End Sub

    Private Sub grdSPTDCOMX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTDCOMX.BeforeRowUpdate
        'If e.Row.IsAddRow Then
        '    e.Row.Cells("CUST_CODE").Value = HFs("CUST_CODE")
        'End If
        If EntryMode = "W" Then
            If e.Row.Cells("WRITE_OFF").Text = "1" Then
                e.Row.Cells("AMT_TO_WOFF").Value = e.Row.Cells("AMT_COMM").Value
            Else
                e.Row.Cells("AMT_TO_WOFF").Value = 0
            End If
        End If
    End Sub

    Private Sub grdSPTDCOMX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTDCOMX.ClickCellButton
        'Dim sql_where As String = ""
        'Select Case grdSPTDCOMX.ActiveCell.Column.Key
        '    Case "SELL_CODE"
        'End Select

        'grdClickCellButton(grdSPTDCOMX, sql_where, False)
    End Sub

    Private Sub grdSPTDCOMX_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSPTDCOMX.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In grdSPTDCOMX.Selected.Rows
        '    If dst.Tables("ARTCUST2").Rows(grow.ListIndex).RowState = DataRowState.Added Then
        '    Else
        '        MsgBox("Cannot Delete Existing Store Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '        Exit For
        '    End If
        'Next
    End Sub

    Private Sub grdSPTDCOMX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTDCOMX.AfterRowActivate
        'If grdSPTDCOMX.ActiveRow.IsAddRow Then
        '    grdSPTDCOMX.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        'Else
        '    grdSPTDCOMX.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        'End If
    End Sub

    Private Sub grdSPTDCOMX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTDCOMX.AfterExitEditMode

        '    With grdSPTDCOMX
        '        Select Case .ActiveCell.Column.Key
        '            Case "CUST_STORE_NO"
        '                If .ActiveCell.Text <> "" Then
        '                    .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
        '                End If
        '        End Select
        '    End With
    End Sub
#End Region

    Sub Load_Open_Accruals()
        Clear_All_Filters(grdSPTDCOMX)
        Fill_Records("SPTDCOMX")

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_DEMO_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")
        Fill_Records("ARTOPENX", New String() {REASON_CODE, ""})
        Sort_grdColumns(grdARTOPENX, "INV_NUM")

        Fill_Records("SPTDPMTX")
        Sort_grdColumns(grdSPTDPMTX, "PYMT_NO".ToLower)
    End Sub

    Private Sub grdSPTDCOMX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTDCOMX.DoubleClickRow
        If (EntryMode = "W") Then
        Else
            If e.Row.IsDataRow Then
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
                Click_Command("View")
            End If
        End If
    End Sub

    Sub Setup_grdSPTDCOMC()
        If grdSPTDCOMC.ActiveRow Is Nothing OrElse Not grdSPTDCOMC.ActiveRow.IsDataRow Then
            grdSPTDCOMS.Visible = False
            grdSPTDCOML.Visible = False
        Else
            Dim OPS_YYYYPP As String = grdSPTDCOMC.ActiveRow.Cells("OPS_YYYYPP").Value
            Dim CUST_CODE As String = grdSPTDCOMC.ActiveRow.Cells("CUST_CODE").Value
            Dim HC_CODE As String = grdSPTDCOMC.ActiveRow.Cells("HC_CODE").Value

            grdSPTDCOMS.Visible = True
            Fill_Records("SPTDCOMS", New String() {OPS_YYYYPP, CUST_CODE, HC_CODE})
            Sort_grdColumns(grdSPTDCOMS, "CUST_STORE_NO")
            grdSPTDCOMS.Text = "Detail by Store - " & OPS_YYYYPP & ":" & HC_CODE

            grdSPTDCOML.Visible = True
            Fill_Records("SPTDCOML", New String() {OPS_YYYYPP, CUST_CODE, HC_CODE})
            Sort_grdColumns(grdSPTDCOML, "COLLECTION_CODE")
            grdSPTDCOML.Text = "Detail by Collection - " & OPS_YYYYPP & ":" & HC_CODE
        End If
    End Sub

    Private Sub grdARTOPENX_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTOPENX.AfterCellUpdate

    End Sub

    Private Sub grdARTOPENX_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdARTOPENX.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdARTOPENX_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTOPENX.ClickCellButton
        If Val(e.Cell.Row.Cells("AMT_TO_APPLY").Value & "") <> 0 Then
            e.Cell.Row.Cells("AMT_TO_APPLY").Value = 0
        Else
            e.Cell.Row.Cells("AMT_TO_APPLY").Value = e.Cell.Row.Cells("INV_BALANCE").Value
        End If
        e.Cell.Row.Update()
    End Sub

    Private Sub grdARTOPENX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTOPENX.DoubleClickRow
        If Not ScreenMode Then
            If e.Row.IsDataRow Then
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
                Click_Command("View")
            End If
        End If
    End Sub

    Private Sub grdARTOPENX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTOPENX.InitializeLayout

    End Sub

#Region "grdSPTDCOMC"

    Private Sub grdSPTDCOMC_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTDCOMC.AfterCellUpdate
        If e.Cell.Column.Key = "AMT_COMM_PAID" Then
            If e.Cell.Value & "" = "" Then
                If e.Cell.Row.Cells("APPLY").Value & "" <> "0" Then
                    e.Cell.Row.Cells("APPLY").Value = "0"
                End If
            Else
                If e.Cell.Row.Cells("APPLY").Value & "" <> "1" Then
                    e.Cell.Row.Cells("APPLY").Value = "1"
                End If
            End If

            e.Cell.Row.Update()

            Dim AMT_COMM_RED_CALC As Decimal = Val(e.Cell.Row.Cells("AMT_COMM_RED_CALC").Value & "")
            Dim AMT_COMM_ADD_CALC As Decimal = Val(e.Cell.Row.Cells("AMT_COMM_ADD_CALC").Value & "")
            Dim r As Int32 = dst.Tables("SPTDCOMC").Rows.Count
            If auto_writeoff Then
            Else
                If (AMT_COMM_RED_CALC <> 0 Or AMT_COMM_ADD_CALC <> 0) And e.Cell.Row.Cells("LEAVE_OPEN").Value & "" <> "1" Then
                    e.Cell.Row.Cells("LEAVE_OPEN").Value = "1"
                ElseIf (AMT_COMM_RED_CALC = 0 And AMT_COMM_ADD_CALC = 0) And e.Cell.Row.Cells("LEAVE_OPEN").Value & "" = "1" Then
                    e.Cell.Row.Cells("LEAVE_OPEN").Value = "0"
                End If
            End If

        End If

        If e.Cell.Column.Key = "APPLY" Then
            If e.Cell.Value & "" = "0" Then
                e.Cell.Row.Cells("AMT_COMM_PAID").Value = DBNull.Value
            Else
                'If e.Cell.Row.Cells("AMT_COMM_PAID").Value & "" = "" Then
                e.Cell.Row.Cells("AMT_COMM_PAID").Value = e.Cell.Row.Cells("AMT_COMM").Value
                'End If
            End If
            e.Cell.Row.Update()
        End If

    End Sub

    Private Sub grdSPTDCOMC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTDCOMC.AfterRowActivate
        Setup_grdSPTDCOMC()
    End Sub
    Private Sub grdSPTDCOMC_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOMC.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSPTDCOMC_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSPTDCOMC.ClickCell
        'If e.Cell IsNot Nothing Then
        '    If e.Cell.DataChanged And e.Cell.Column.Key = "APPLY" Then
        '        ' e.Cell.Row.Update()
        '        If e.Cell.Value = "0" Then
        '            e.Cell.Row.Cells("AMT_COMM_PAID").Value = DBNull.Value
        '        Else
        '            e.Cell.Row.Cells("AMT_COMM_PAID").Value = e.Cell.Row.Cells("AMT_COMM").Value
        '        End If
        '        e.Cell.Row.Update()
        '    End If
        'End If
    End Sub

    Private Sub grdSPTDCOMC_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTDCOMC.ClickCellButton
        'If Val(e.Cell.Row.Cells("AMT_COMM_PAID").Value & "") <> 0 Then
        '    e.Cell.Row.Cells("AMT_COMM_PAID").Value = 0
        'Else
        '    e.Cell.Row.Cells("AMT_COMM_PAID").Value = e.Cell.Row.Cells("AMT_COMM").Value
        'End If
        'e.Cell.Row.Update()
    End Sub

    Private Sub grdSPTDCOMC_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTDCOMC.InitializeRow

        With e.Row.Cells("AMT_COMM_PAID")
            If Val(e.Row.Cells("AMT_COMM_ADD").Value & "") <> 0 Or Val(e.Row.Cells("AMT_COMM_RED").Value & "") <> 0 Then
                .Appearance.ForeColor = Drawing.Color.Red
                .ToolTipText = "More or Less Expense was applied on this Payment"
            Else
                .Appearance.ForeColor = Drawing.Color.Empty
                .ToolTipText = ""
            End If
        End With
    End Sub

    Private Sub grdSPTDCOMC_MouseClick(sender As Object, e As MouseEventArgs) Handles grdSPTDCOMC.MouseClick
        'If grdSPTDCOMC.ActiveRow IsNot Nothing AndAlso grdSPTDCOMC.ActiveCell IsNot Nothing Then
        '    If grdSPTDCOMC.ActiveCell.DataChanged And grdSPTDCOMC.ActiveCell.Column.Key = "APPLY" Then
        '        grdSPTDCOMC.ActiveRow.Update()
        '        If grdSPTDCOMC.ActiveCell.Value = "0" Then
        '            grdSPTDCOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = DBNull.Value
        '        Else
        '            grdSPTDCOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = grdSPTDCOMC.ActiveRow.Cells("AMT_COMM").Value
        '        End If
        '        grdSPTDCOMC.ActiveRow.Update()
        '    End If
        'End If
    End Sub

    Private Sub grdSPTDCOMC_MouseUp(sender As Object, e As MouseEventArgs) Handles grdSPTDCOMC.MouseUp
        If grdSPTDCOMC.ActiveRow IsNot Nothing AndAlso grdSPTDCOMC.ActiveCell IsNot Nothing Then
            If grdSPTDCOMC.ActiveCell.DataChanged And grdSPTDCOMC.ActiveCell.Column.Key = "APPLY" Then
                grdSPTDCOMC.ActiveRow.Update()
                'If grdSPTDCOMC.ActiveCell.Value = "0" Then
                '    grdSPTDCOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = DBNull.Value
                'Else
                '    grdSPTDCOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = grdSPTDCOMC.ActiveRow.Cells("AMT_COMM").Value
                'End If
                'grdSPTDCOMC.ActiveRow.Update()
            End If
        End If
    End Sub

#End Region
    Sub Display_Totals()
        Dim CREDITED As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(AMT_COMM_PAID)", "") & "")
        Dim DEDUCTED As Decimal = Val(dst.Tables("ARTOPENX").Compute("SUM(AMT_TO_APPLY)", "") & "")
        Dim AMT_COMM As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(AMT_COMM)", "APPLY='1'") & "")
        Dim AMT_COMM_ADD As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(AMT_COMM_ADD)", "APPLY='1'") & "")
        Dim AMT_COMM_RED As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(AMT_COMM_RED)", "APPLY='1'") & "")
        Dim BALANCE As Decimal = Val(dst.Tables("SPTDCOMC").Compute("SUM(BALANCE)", "APPLY='1'") & "")


        Dim DEMO_COMM_OFFSET As Decimal = AMT_COMM - BALANCE
        Dim DEMO_COMM_ADD_EXP As Decimal = AMT_COMM_ADD - AMT_COMM_RED

        'rowSPTDPMT1 IsNot Nothing AndAlso

        If ScreenMode AndAlso EntryMode <> "L" AndAlso (auto_writeoff Or rowSPTDPMT1.Item("REVERSED_PYMT_NO") & "" <> "") Then
            'If EntryMode = "L" Then Stop
            'If rowSPTDPMT1 Is Nothing Then Stop
            'If auto_writeoff Then Stop

            DEMO_COMM_OFFSET = Val(dst.Tables("SPTDCOMC").Compute("SUM(AMT_COMM_OFFSET)", "APPLY='1'") & "")
            DEMO_COMM_ADD_EXP = CREDITED - DEMO_COMM_OFFSET
        End If

        numCREDITED.Value = CREDITED
        numDEDUCTED.Value = DEDUCTED
        numDIFFERENCE.Value = DEDUCTED - CREDITED
        numDEMO_COMM_OFFSET.Value = DEMO_COMM_OFFSET
        numDEMO_COMM_ADD_EXP.Value = DEMO_COMM_ADD_EXP ' AMT_COMM_ADD - AMT_COMM_RED
    End Sub

    Private Sub grdARTOPENX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPENX.InitializeRow
        e.Row.Cells("PAY").Value = "->"
    End Sub

    Private Sub grdSPTDPMTX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTDPMTX.DoubleClickRow

        If e.Row.IsDataRow Then
            PYMT_NO = e.Row.Cells("PYMT_NO").Value
            If PYMT_NO <> "" Then
                rowSPTDPMT1 = LookUp("SPTDPMT1", PYMT_NO)
                If rowSPTDPMT1 IsNot Nothing Then
                    CUST_CODE = rowSPTDPMT1.Item("CUST_CODE")
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                    Click_Command("View Payment")
                End If
            End If
        End If
    End Sub

    Private Sub btnWriteOff_Click(sender As Object, e As EventArgs) Handles btnWriteOff.Click
        Dim YYYY As String = "2015"

        If MsgBox("OK to Write Off " & YYYY, MsgBoxStyle.OkCancel, "Verification") <> MsgBoxResult.Ok Then Exit Sub

        auto_writeoff = True

        ASCMAIN1.sql = "Select CUST_CODE, COUNT (*) RECS, SUM (AMT_COMM) AMT_COMM from SPTDCOMC where OPS_YYYYPP LIKE '" & YYYY & "%' AND PYMT_NO IS NULL group by CUST_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable().Select("")

            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim RECS As Int64 = Val(row.Item("RECS") & "")
            Dim AMT_COMM As Decimal = Val(row.Item("AMT_COMM") & "")
            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Click_Command("View")
            Click_Command("Enter Payment")

            'Absx1.txtFor("PYMT_REF_NO").Text = "Auto Write-Off " & YYYY
            'Absx1.dteFor("PYMT_REF_DATE").Value = Now.Date

            ' THESE VALUES DON'T GET PUT INTO DATATABLE WHEN PROGRAMATICALLY SET
            rowSPTDPMT1.Item("PYMT_REF_DATE") = Now.Date ' Absx1.dteFor("PYMT_REF_DATE").Value
            rowSPTDPMT1.Item("PYMT_REF_NO") = "WOff " & YYYY ' Absx1.txtFor("PYMT_REF_NO").Text

            For Each grow As UltraWinGrid.UltraGridRow In grdSPTDCOMC.Rows
                Dim OPS_YYYYPP As String = grow.Cells("OPS_YYYYPP").Value
                If OPS_YYYYPP.StartsWith(YYYY) Then
                    grow.Cells("APPLY").Value = "1"
                    grow.Cells("AMT_COMM_PAID").Value = 0
                    grow.Update()
                End If
            Next

            If Val(Absx1.numFor("DEMO_COMM_OFFSET").Value & "") <> AMT_COMM Then
                Stop
            End If

            Click_Command("Update")
            If Not auto_writeoff Then
                Exit For
            End If
        Next

        auto_writeoff = False
        MsgBox("Write Off " & YYYY & " is Complete", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Reverse_Record()

        Dim PYMT_REF_DATE As Date = Absx1.dteFor("PYMT_REF_DATE").Value
        Dim PYMT_REF_NO As String = Absx1.txtFor("PYMT_REF_NO").Text

        Dim PYMT_CTL_NO As String = rowSPTDPMT1.Item("PYMT_CTL_NO")
        Dim PYMT_CTL_NO_R As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")

        Dim rowSPTDPMT1_R As DataRow = dst.Tables("SPTDPMT1").NewRow
        rowSPTDPMT1_R.ItemArray = rowSPTDPMT1.ItemArray
        Dim PYMT_NO_R As String = ASCMAIN1.Next_Control_No("SPTDPMT1.PYMT_NO")
        With rowSPTDPMT1_R
            .Item("PYMT_NO") = PYMT_NO_R
            .Item("REVERSED_PYMT_NO") = PYMT_NO

            .Item("PYMT_DATE") = DATETIME_STAMP.Date
            .Item("PYMT_REF_DATE") = DATETIME_STAMP.Date

            For Each C As String In New String() {"PYMT_REF_AMT", "DEMO_COMM_OFFSET", "DEMO_COMM_CREDITED", "DEMO_COMM_APPLIED", "DEMO_COMM_ADD_EXP"}
                .Item(C) = -1 * Val(.Item(C) & "")
            Next

            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID

            .Item("JOURNAL_IND") = DBNull.Value
            .Item("JOURNAL_XNO") = DBNull.Value

            .Item("PYMT_CTL_NO") = PYMT_CTL_NO_R

        End With
        With rowSPTDPMT1
            .Item("REVERSED_BY_PYMT_NO") = PYMT_NO_R
        End With

        dst.Tables("SPTDPMT1").Rows.Add(rowSPTDPMT1_R)

        Fill_Records("SPTDPMT2", PYMT_NO)
        For Each rowSPTDPMT2 As DataRow In dst.Tables("SPTDPMT2").Select("")
            With rowSPTDPMT2
                .Item("PYMT_NO") = PYMT_NO_R
                .Item("AMT_COMM_ADJ") = -1 * Val(.Item("AMT_COMM_ADJ") & "")
                .AcceptChanges()
                .SetAdded()
            End With
        Next


        'For Each rowSPTDCOMC As DataRow In dst.Tables("SPTDCOMC").Select("")
        '    With rowSPTDCOMC
        '        If (.Item("ACC_CTL_NO_ORIG") & "") = "" Then .Item("ACC_CTL_NO_ORIG") = .Item("ACC_CTL_NO")
        '        .Item("ACC_CTL_NO") = ASCMAIN1.Next_Control_No("SPTDCOMC.ACC_CTL_NO")
        '        .Item("AMT_COMM") = Val(.Item("AMT_COMM") & "") - Val(.Item("AMT_COMM_OFFSET") & "")
        '        .Item("AMT_COMM_ADJ") = -1 * Val(.Item("AMT_COMM_ADJ") & "")
        '        .Item("AMT_COMM_OFFSET") = -1 * Val(.Item("AMT_COMM_OFFSET") & "")
        '        .Item("AMT_COMM_PAID") = -1 * Val(.Item("AMT_COMM_PAID") & "")
        '        '.Item("OPS_YYYYPP_PAID") = ASCMAIN1.CYP
        '        '.Item("PYMT_NO") = PYMT_NO_R
        '        .Item("PYMT_NO_ORIG") = PYMT_NO
        '    End With
        'Next


        ' iterate through original C records to create new C records to connect to the reversal payment

        For Each rowSPTDCOMC_orig As DataRow In dst.Tables("SPTDCOMC").Select("APPLY = '1'", "CUST_CODE,OPS_YYYYPP")
            Dim ACC_CTL_NO_orig As String = rowSPTDCOMC_orig.Item("ACC_CTL_NO_ORIG") & ""
            If ACC_CTL_NO_orig = "" Then ACC_CTL_NO_orig = rowSPTDCOMC_orig.Item("ACC_CTL_NO") & ""

            Dim ACC_CTL_NO As String = Replicate_Accrual(rowSPTDCOMC_orig.Item("ACC_CTL_NO"), True)
            Dim rowSPTDCOMC_R As DataRow = dst.Tables("SPTDCOMC").Rows.Find(ACC_CTL_NO)

            With rowSPTDCOMC_R
                .Item("ACC_CTL_NO_ORIG") = ACC_CTL_NO_orig
                .Item("JOURNAL_XNO") = DBNull.Value
                .Item("JOURNAL_IND") = DBNull.Value
                .Item("AMT_COMM_OFFSET") = -1 * Val(rowSPTDCOMC_orig.Item("AMT_COMM_OFFSET"))
                .Item("AMT_SOLD_CLAIMED") = -1 * Val(rowSPTDCOMC_orig.Item("AMT_SOLD_CLAIMED"))
                .Item("AMT_COMM_PAID") = -1 * Val(rowSPTDCOMC_orig.Item("AMT_COMM_PAID"))
                .Item("AMT_COMM_ADJ") = -1 * Val(rowSPTDCOMC_orig.Item("AMT_COMM_ADJ"))

                .Item("OPS_YYYYPP_PAID") = ASCMAIN1.CYP
                .Item("PYMT_NO") = PYMT_NO_R

                .Item("APPLY") = "1"

                'If Val(.Item("AMT_COMM_ADJ") & "") = 0 Then
                .Item("LEAVE_OPEN") = "1"
                'Else
                '    .Item("LEAVE_OPEN") = "0"
                'End If
            End With
        Next

        ' Transfer ACC_CTL_NOs to ACC_CTL_NOs_R

        Dim ACC_CTL_NOs_R As New List(Of String)
        For Each ACC_CTL_NO As String In ACC_CTL_NOs
            ACC_CTL_NOs_R.Add(ACC_CTL_NO)
        Next

        Dim OPS_YYYYPP As String = ""

        Dim DEMO_COMM_OFFSET As Decimal = 0
        Dim DEMO_COMM_CREDITED As Decimal = 0

        ' iterate through new C records to record activity into B 

        For Each ACC_CTL_NO As String In ACC_CTL_NOs_R
            Dim rowSPTDCOMC As DataRow = dst.Tables("SPTDCOMC").Rows.Find(ACC_CTL_NO)

            If OPS_YYYYPP <> rowSPTDCOMC.Item("OPS_YYYYPP") Then
                OPS_YYYYPP = rowSPTDCOMC.Item("OPS_YYYYPP")
                Fill_Records("SPTDCOMB", New String() {CUST_CODE, OPS_YYYYPP}, False)
            End If

            Dim HC_CODE As String = rowSPTDCOMC.Item("HC_CODE")
            Dim AMT_COMM_total As Decimal = Val(rowSPTDCOMC.Item("AMT_COMM") & "")
            Dim AMT_COMM_PAID_total As Decimal = Val(rowSPTDCOMC.Item("AMT_COMM_PAID") & "")

            DEMO_COMM_CREDITED += AMT_COMM_PAID_total

            If rowSPTDCOMC.Item("LEAVE_OPEN") & "" = "1" Then
                DEMO_COMM_OFFSET += AMT_COMM_PAID_total
                'rowSPTDCOMC.Item("AMT_COMM_ADJ") = 0
                'rowSPTDCOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_PAID_total
            Else
                DEMO_COMM_OFFSET += AMT_COMM_total
                'rowSPTDCOMC.Item("AMT_COMM_ADJ") = AMT_COMM_PAID_total - AMT_COMM_total
                'rowSPTDCOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_total
            End If

            Dim AMT_COMM_PAID_spread As Decimal = 0

            Dim sqlw As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and CUST_CODE = '" & CUST_CODE & "' and HC_CODE = '" & HC_CODE & "'"
            Dim AMT_COMM_all As Decimal = Val(dst.Tables("SPTDCOMB").Compute("SUM(AMT_COMM)", sqlw) & "")

            Dim rows() As DataRow = dst.Tables("SPTDCOMB").Select(sqlw, "AMT_COMM")
            Dim r As Integer = 0

            For Each rowSPTDCOMB As DataRow In rows

                Dim AMT_COMM As Decimal = Val(rowSPTDCOMB.Item("AMT_COMM") & "")
                Dim AMT_COMM_PAID As Decimal = Val(rowSPTDCOMB.Item("AMT_COMM_PAID") & "")

                Dim F As Decimal = 1
                If AMT_COMM_all <> 0 Then F = AMT_COMM / AMT_COMM_all

                Dim AMT_COMM_PAID_now As Decimal = System.Math.Round(AMT_COMM_PAID_total * F, 2)
                AMT_COMM_PAID_spread += AMT_COMM_PAID_now
                AMT_COMM_PAID += AMT_COMM_PAID_now

                r += 1
                If r = rows.Length And AMT_COMM_PAID_total <> AMT_COMM_PAID_spread Then
                    AMT_COMM_PAID += AMT_COMM_PAID_total - AMT_COMM_PAID_spread
                End If

                rowSPTDCOMB.Item("AMT_COMM_PAID") = AMT_COMM_PAID
            Next

            ' record new C record with balance open - note that all reversals have LEAVE_OPEN = 1 and BALANCE <> 0
            Dim ACC_CTL_NO_orig As String = rowSPTDCOMC.Item("ACC_CTL_NO_ORIG") & ""
            If ACC_CTL_NO_orig = "" Then ACC_CTL_NO_orig = rowSPTDCOMC.Item("ACC_CTL_NO") & ""
            Dim ACC_CTL_NO_new As String = Replicate_Accrual(rowSPTDCOMC.Item("ACC_CTL_NO"), True)
            Dim rowSPTDCOMC_new As DataRow = dst.Tables("SPTDCOMC").Rows.Find(ACC_CTL_NO_new)

            rowSPTDCOMC_new.Item("ACC_CTL_NO_ORIG") = ACC_CTL_NO_orig
            rowSPTDCOMC_new.Item("PYMT_NO") = DBNull.Value
            If Val(rowSPTDCOMC_new.Item("AMT_COMM_ADJ") & "") <> 0 Then
                rowSPTDCOMC_new.Item("AMT_COMM") = Val(rowSPTDCOMC_new.Item("AMT_COMM") & "") + Val(rowSPTDCOMC_new.Item("AMT_COMM_ADJ") & "")
                rowSPTDCOMC_new.Item("AMT_COMM_ADJ") = 0
            End If
        Next

        Dim INV_TOTAL_AMOUNT As Decimal = -1 * Val(Absx1.numFor("CREDITED").Value & "")
        Dim DEDUCTED As Decimal = -1 * Val(Absx1.numFor("DEDUCTED").Value & "")
        Dim INV_DATE As Date = Now.Date

        BeginTrans()

        Dim PYMT_BATCH_NO As String = rowSPTDPMT1.Item("PYMT_BATCH_NO") ' ORIGINAL PYMT_BATCH_NO - NEEDED TO RETREIVE CBS APPLIED

        ' PUT THE ARTOPENX RECORDS INTO ARTOPEN1
        Dim sqlwAR As String = $"(INV_TYPE, INV_NUM) In (Select INV_TYPE, INV_NUM FROM ARTPYMT3 where PYMT_BATCH_NO = '{PYMT_BATCH_NO}' and INV_NUM <> '{PYMT_CTL_NO}')"

        ASCMAIN1.sql = $"Update ARTOPENX Set OPS_YYYYPP_F = NULL where {sqlwAR}"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Insert into ARTOPEN1 Select * from ARTOPENX where {sqlwAR}"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Delete from ARTOPENX where {sqlwAR}"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = sqlARTOPENX
        ASCMAIN1.sql &= $" from ARTOPEN1 WHERE {sqlwAR}"
        Fill_Records("ARTOPENX", ,, ASCMAIN1.sql)

        ASCMAIN1.sql = $"SELECT INV_TYPE, INV_NUM, INV_PMT FROM ARTPYMT3 WHERE PYMT_BATCH_NO = '{PYMT_BATCH_NO}' and INV_NUM <> '{PYMT_CTL_NO}'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim INV_TYPE As String = row.Item("INV_TYPE")
            Dim INV_NUM As String = row.Item("INV_NUM")
            Dim INV_PMT As String = Val(row.Item("INV_PMT") & "")
            Dim rowARTOPENX As DataRow = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_CODE, INV_TYPE, INV_NUM})
            rowARTOPENX.Item("AMT_TO_APPLY") = -1 * INV_PMT
        Next

        Dim PYMT_BATCH_NO_R As String = Setup_ARTPYMTx(PYMT_CTL_NO_R, PYMT_REF_DATE, PYMT_REF_NO, INV_DATE, INV_TOTAL_AMOUNT, DEDUCTED)
        rowSPTDPMT1_R.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_R

        Update_Record_TDA("SPTDCOMC")
        Update_Record_TDA("SPTDCOMB")

        Update_Record_TDA("SPTDPMT1")
        Update_Record_TDA("SPTDPMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTOPEN1")

        ASCMAIN1.Record_Event("SPTDPMT1", PYMT_NO_R, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "P", "Reverse Payment Entered", "")
        ASCMAIN1.Record_Event("SPTDPMT1", PYMT_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "P", "Payment Reversed", "")

        CommitTrans($"Demo Payment {PYMT_NO} has been Reversed by Payment {PYMT_NO_R}")

    End Sub


    Function Setup_ARTPYMTx(PYMT_CTL_NO As String, PYMT_REF_DATE As Date, PYMT_REF_NO As String, INV_DATE As Date, INV_TOTAL_AMOUNT As Decimal, DEDUCTED As Decimal) As String

        Dim PYMT_BATCH_NO As String = ""

        Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_DEMO_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)

        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
        Dim TERM_CODE As String = rowSOTTYPE1.Item("TERM_CODE")
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        With rowARTOPEN1
            .Item("CUST_CODE") = CUST_BILL_TO_CUST
            .Item("INV_TYPE") = "C"
            .Item("INV_NUM") = PYMT_CTL_NO
            .Item("INV_DATE") = INV_DATE
            .Item("POST_CODE") = POST_CODE
            .Item("TERM_CODE") = TERM_CODE
            .Item("INV_DUE_DATE") = INV_DATE
            .Item("INV_CUST_PO") = PYMT_REF_NO
            .Item("INV_MISC_CHG") = -1 * INV_TOTAL_AMOUNT
            .Item("INV_TOTAL_AMOUNT") = -1 * INV_TOTAL_AMOUNT
            .Item("INV_BALANCE") = -1 * INV_TOTAL_AMOUNT
            .Item("CUST_CODE_SO") = CUST_CODE
            .Item("REASON_CODE") = REASON_CODE
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP

            .Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE

            .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            .Item("SALES_DIVISION_CODE") = ""

            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            .Item("INV_SALES_CURR") = .Item("INV_SALES")
            .Item("INV_DISC_CURR") = .Item("INV_DISC")
            .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
            .Item("INV_STAX_CURR") = .Item("INV_STAX")
            .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")
            .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
            .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")

            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1
        End With

        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        If dst.Tables("ARTOPENX").Select("AMT_TO_APPLY <> 0").Length <> 0 Then

            PYMT_BATCH_NO = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")

            Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
            With rowARTPYMT1
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_DATE") = INV_DATE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("STATUS") = "1"
                .Item("PYMT_APPL_ONLY") = "1"
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1
                .Item("PYMT_SOURCE") = "D"
            End With
            dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

            Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
            With rowARTPYMT2
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_LNO") = 1
                .Item("CUST_CODE") = CUST_BILL_TO_CUST
                .Item("CUST_NAME") = rowARTCUST1_BT.Item("CUST_NAME")
                .Item("CUST_PYMT_REF_NO") = PYMT_REF_NO
                .Item("CUST_PYMT_REF_DATE") = PYMT_REF_DATE
                .Item("CUST_PYMT_AMT") = 0
                .Item("PYMT_STATUS") = "2"
                .Item("CUST_PYMT_AMT_CURR") = 0
                .Item("PYMT_NOTE") = "Demo Payment"
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With
            dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

            dst.Tables("ARTPYMT3").Rows.Clear()

            ' Record Credit in ARTPYMT3
            Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, PYMT_BATCH_NO, 0, -1 * DEDUCTED, PYMT_REF_NO, PYMT_REF_DATE)

            ' Record Payments against Chargebacks
            Dim PYMT_BATCH_ILNO As Integer = 0
            For Each rowARTOPENX As DataRow In dst.Tables("ARTOPENX").Select("AMT_TO_APPLY <> 0", "INV_NUM")
                Dim CUST_CODE As String = rowARTOPENX.Item("CUST_CODE")
                Dim INV_TYPE As String = rowARTOPENX.Item("INV_TYPE")
                Dim INV_NUM As String = rowARTOPENX.Item("INV_NUM")
                Dim INV_PMT As Decimal = Val(rowARTOPENX.Item("AMT_TO_APPLY") & "")
                rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {CUST_CODE, INV_TYPE, INV_NUM}, False, False)
                PYMT_BATCH_ILNO += 1
                Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, PYMT_BATCH_NO, PYMT_BATCH_ILNO, INV_PMT, PYMT_REF_NO, PYMT_REF_DATE)
            Next
        End If

        Return PYMT_BATCH_NO
    End Function
End Class