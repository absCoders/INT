Public Class SPFAPMT1

    Dim rowSPTAPMT1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim PYMT_NO As String
    Dim PYMT_BATCH_NO As String
    Dim sqlARTOPENX As String
    Dim ACC_CTL_NOs As New List(Of String)
    Dim REASON_CODE As String = ""

    Dim CUST_BILL_TO_CUST As String
    Dim rowARTCUST1_BT As DataRow

    Dim SPTAPMTI As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")

        With dst
            ASCMAIN1.sql = "Select * from SPTACOMC where PYMT_NO is Null"
            Create_TDA(.Tables.Add, "SPTACOMX", "**", 0, False)

            sqlARTOPENX = "Select ARTOPEN1.CUST_CODE, ARTOPEN1.INV_TYPE, ARTOPEN1.INV_NUM" & vbCrLf _
                & ", ARTOPEN1.INV_DATE, ARTOPEN1.INV_CUST_PO, ARTOPEN1.INV_TOTAL_AMOUNT, ARTOPEN1.INV_BALANCE" & vbCrLf _
                & ", ARTOPEN1.INIT_OPER, ARTOPEN1.INIT_DATE, ARTOPEN1.REASON_CODE, ARTOPEN1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ARTOPEN1.INV_NOTES, ARTOPEN1.ORDR_TYPE_CODE, ARTOPEN1.OPS_YYYYPP" & vbCrLf

            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = sqlARTOPENX _
                    & " from ARTOPEN1 where (INV_TYPE = 'B' or INV_TYPE = 'O') and CUST_CODE = NVL(:PARM2,CUST_CODE)"
                Create_TDA(.Tables.Add, "ARTOPENX", "**", 0, False, "V")
            Else
                ASCMAIN1.sql = sqlARTOPENX _
                    & " from ARTOPEN1 where (INV_TYPE = 'B' or INV_TYPE = 'O') and REASON_CODE = :PARM1 and CUST_CODE = NVL(:PARM2,CUST_CODE)"
                Create_TDA(.Tables.Add, "ARTOPENX", "**", 0, False, "VV")
            End If

            With .Tables("ARTOPENX")
                .Columns.Add("AMT_TO_APPLY", GetType(System.Decimal))
                .Columns.Add("INV_BALANCE_NEW", GetType(System.Decimal), "ISNULL(INV_BALANCE,0)-ISNULL(AMT_TO_APPLY,0)")
            End With

            'ASCMAIN1.sql = "Select ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, ARTPYMT2.CUST_CODE, ARTPYMT2.CUST_NAME, ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE, ARTPYMT2.PYMT_NOTE, ARTPYMT2.INIT_DATE, ARTPYMT2.INIT_OPER from ARTPYMT2,ARTPYMT1" & vbCrLf _
            '    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            '    & "   and ARTPYMT1.PYMT_SOURCE = 'D'"
            ASCMAIN1.sql = "Select * from SPTAPMT1"
            Create_TDA(.Tables.Add, "SPTAPMTX", "**", 0, False)

            ASCMAIN1.sql = "Select OPS_YYYYPP, ASP_CODE, CUST_CODE, CUST_STORE_NO, INV_NO" & vbCrLf _
                & ", COUNT (*) COLLECTIONS" & vbCrLf _
                & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
                & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
                & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
                & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
                & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
                & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
                & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
                & " from SPTACOMB" & vbCrLf _
                & " where OPS_YYYYPP = :PARM1 and CUST_CODE = :PARM2 and HC_CODE = NVL(:PARM3,HC_CODE) and BRAND_CODE = :PARM4 and ASP_CODE = :PARM5 and INV_NO = :PARM6 and ACC_CTL_NO_ACCRUAL = :PARM7" & vbCrLf _
                & " group by OPS_YYYYPP, ASP_CODE, CUST_CODE, CUST_STORE_NO, INV_NO"
            Create_TDA(.Tables.Add, "SPTACOMS", "**", 0, False, "VVVVVVV", 5)

            ASCMAIN1.sql = "Select OPS_YYYYPP, ASP_CODE, CUST_CODE, COLLECTION_CODE, INV_NO" & vbCrLf _
                & ", COUNT (*) STORES" & vbCrLf _
                & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
                & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
                & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
                & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
                & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
                & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
                & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
                & " from SPTACOMB" & vbCrLf _
                & " where OPS_YYYYPP = :PARM1 and CUST_CODE = :PARM2 and HC_CODE = NVL(:PARM3,HC_CODE) and BRAND_CODE = :PARM4 and ASP_CODE = :PARM5 and INV_NO = :PARM6 and ACC_CTL_NO_ACCRUAL = :PARM7" & vbCrLf _
                & " group by OPS_YYYYPP, ASP_CODE, CUST_CODE, COLLECTION_CODE, INV_NO"
            Create_TDA(.Tables.Add, "SPTACOML", "**", 0, False, "VVVVVVV", 5)
            ', SPTACOM1.ASP_COMM_PCT
            ASCMAIN1.sql = "Select Distinct ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.CUST_BILL_TO_CUST, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from ARTCUST1,SPTACOM1 where SPTACOM1.CUST_CODE (+) = ARTCUST1.CUST_CODE and ARTCUST1.CUST_CODE =:PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from SPTACOMC where CUST_CODE = :PARM1 and OPS_YYYYPP_PAID is Null"
            Create_TDA(.Tables.Add, "SPTACOMC", "**", 0, True, "V")
            With .Tables("SPTACOMC")
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

            ASCMAIN1.sql = "Select * from SPTACOMB where CUST_CODE = :PARM1 and OPS_YYYYPP = :PARM2"
            Create_TDA(.Tables.Add, "SPTACOMB", "**", 0, True, "VV")

            ASCMAIN1.sql = "Select * from SPTACOMC where CUST_CODE = :PARM1 and OPS_YYYYPP_PAID is Not Null"
            Create_TDA(.Tables.Add, "SPTACOMP", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SPTAPMT1", "*")
            Create_TDA(.Tables.Add, "SPTAPMT2", "*")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")

            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
            .Tables("ARTPYMT3").Columns.Add("SELECTED")
            .Tables("ARTPYMT3").Columns("SELECTED").DefaultValue = "0"
            .Tables("ARTPYMT3").Columns.Add("CUST_CODE")


            ASCMAIN1.sql = "Select INV_NUM from ARTOPEN1 where ROWNUM < 1"
            SPTAPMTI = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SPTAPMTI & " Add Primary Key (INV_NUM)")
            Create_TDA(.Tables.Add, SPTAPMTI, "*")

            '& ", COUNT (*) INVS, SUM (ARTPYMT5.GL_DIST_AMT) DED, SUM (SOTINVH1.INV_SALES) SLS" _
            ASCMAIN1.sql = "" _
                & "Select SOTINVH2.ORDR_YYYYPP_UPDATED, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * .115) COM" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * .01) ADVBROAD" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * .08) ADVCOOP" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * .02) DIF" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * .005) TESTERS" & vbCrLf _
                & " from " & SPTAPMTI & " SPTAPMTI,SOTINVH1,SOTINVH2,ICTITEM1,ICTCOLL1,ARTPYMT5" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ARTPYMT5.CHARGEBACK_NO = SPTAPMTI.INV_NUM" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = ARTPYMT5.CUST_REFERENCE" & vbCrLf _
                & " group by SOTINVH2.ORDR_YYYYPP_UPDATED, ICTCOLL1.BRAND_CODE"
            Create_TDA(.Tables.Add, "SPTAPMTY", "**", 0, False, "", 2)
            With .Tables("SPTAPMTY").Columns
                .Add("PAID_ADVBROAD", GetType(System.Decimal))
                .Add("PAID_ADVCOOP", GetType(System.Decimal))
                .Add("PAID_DIF", GetType(System.Decimal))
                .Add("PAID_TESTERS", GetType(System.Decimal))
                .Add("UNPAID_ADVBROAD", GetType(System.Decimal), "ISNULL(ADVBROAD,0)-ISNULL(PAID_ADVBROAD,0)")
                .Add("UNPAID_ADVCOOP", GetType(System.Decimal), "ISNULL(ADVCOOP,0)-ISNULL(PAID_ADVCOOP,0)")
                .Add("UNPAID_DIF", GetType(System.Decimal), "ISNULL(DIF,0)-ISNULL(PAID_DIF,0)")
                .Add("UNPAID_TESTERS", GetType(System.Decimal), "ISNULL(TESTERS,0)-ISNULL(PAID_TESTERS,0)")
            End With

        End With

        grdSPTACOMC.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide

        grdSPTACOMX.DataSource = dst.Tables("SPTACOMX")
        grdSPTACOMC.DataSource = dst.Tables("SPTACOMC")
        grdSPTACOMP.DataSource = dst.Tables("SPTACOMP")

        grdARTOPENX.DataSource = dst.Tables("ARTOPENX")
        grdSPTAPMTX.DataSource = dst.Tables("SPTAPMTX")

        grdSPTACOMS.DataSource = dst.Tables("SPTACOMS")
        grdSPTACOML.DataSource = dst.Tables("SPTACOML")

        grdSPTAPMTY.DataSource = dst.Tables("SPTAPMTY")

        grd_Appearance_LightGray(grdSPTACOMX)
        grd_Appearance_LightGray(grdSPTACOMC)
        grd_Appearance_LightGray(grdSPTACOMP)
        grd_Appearance_LightGray(grdARTOPENX)
        grd_Appearance_LightGray(grdSPTAPMTX)
        grd_Appearance_LightGray(grdSPTACOMS)
        grd_Appearance_LightGray(grdSPTACOML)

        Create_Summary(grdSPTACOMX, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTACOMX, "AMT_COMM")

        Create_Summary(grdSPTACOMC, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTACOMC, New String() {"QTY_SOLD", "AMT_SOLD", "AMT_COMM", "AMT_SOLD_CLAIMED", "AMT_COMM_PAID", "APPLY", "AMT_COMM_ADD", "AMT_COMM_RED", "BALANCE"})

        Create_Summary(grdSPTACOMP, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTACOMP, New String() {"QTY_SOLD", "AMT_SOLD", "AMT_COMM", "AMT_SOLD_CLAIMED", "AMT_COMM_PAID"})

        Create_Summary(grdARTOPENX, "INV_NUM", "Count")
        Create_Summary(grdARTOPENX, New String() {"INV_BALANCE", "AMT_TO_APPLY", "INV_BALANCE_NEW"})

        Create_Summary(grdSPTACOMS, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTACOMS, New String() {"QTY_SOLD", "AMT_SOLD", "QTY_EOW", "AMT_EOW", "AMT_COMM"})

        Create_Summary(grdSPTACOML, "OPS_YYYYPP", "Count")
        Create_Summary(grdSPTACOML, New String() {"QTY_SOLD", "AMT_SOLD", "QTY_EOW", "AMT_EOW", "AMT_COMM"})

        Create_Summary(grdSPTAPMTX, "PYMT_NO", "Count")
        Create_Summary(grdSPTAPMTX, New String() {"PYMT_REF_AMT", "ASP_COMM_OFFSET", "ASP_COMM_CREDITED", "ASP_COMM_APPLIED", "ASP_COMM_ADD_EXP"})

        Create_Summary(grdSPTAPMTY, "ORDR_YYYYPP_UPDATED", "Count")
        Create_Summary(grdSPTAPMTY, New String() {"COM", _
                                                  "ADVBROAD", "ADVCOOP", "DIF", "TESTERS", _
                                                  "PAID_ADVBROAD", "PAID_ADVCOOP", "PAID_DIF", "PAID_TESTERS",
                                                  "UNPAID_ADVBROAD", "UNPAID_ADVCOOP", "UNPAID_DIF", "UNPAID_TESTERS"})

        Show_Filter(grdSPTACOMX, True)
        grdSPTACOMX.DisplayLayout.GroupByBox.Hidden = False

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTACOMX, grdSPTACOMC, grdSPTACOMP, grdSPTACOMS, grdSPTACOML}

            With grd.DisplayLayout.Bands(0)
                .Columns("QTY_SOLD").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns("AMT_SOLD").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns("AMT_COMM").Header.Appearance.BackColor2 = Drawing.Color.Gold
                If grd.Name = "grdSPTACOMS" Or grd.Name = "grdSPTACOML" Then
                    .Columns("QTY_EOW").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    .Columns("AMT_EOW").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
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

        With grdSPTACOMC.DisplayLayout.Bands(0)
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

        If ASCMAIN1.CLIENT = "AHA" Then
            REASON_CODE = "CM"
            chkShowRCOnly.Text = "Show RC " & REASON_CODE & " Only"
            chkShowRCOnly.Checked = True
            chkShowRCOnly.Visible = True
        Else
            chkShowRCOnly.Checked = False
            chkShowRCOnly.Visible = False
        End If

        SplitContainer1.SplitterDistance = SplitContainer1.Height / 2
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Validate_Code("CUST_CODE")

            Case "Enter Payment"
                If Not ASCMAIN1.Logical_Lock("SPTCOMM1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SPTACOM1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Then Exit Sub

            Case "Update"

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

                    If System.Math.Round(AMT_TO_APPLY_total, 2) > System.Math.Round(CREDITED, 2) Then
                        EMsg &= vbCr & "You cannot apply more than the total Credit Value (" & Format(CREDITED, "$#,##0.00") & ") to Chargebacks"
                    End If

                    Dim PYMT_REF_AMT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")
                    If System.Math.Round(PYMT_REF_AMT, 2) <> System.Math.Round(CREDITED, 2) Then
                        EMsg &= vbCr & "Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ") does not equal Total Credit distribution (" & Format(CREDITED, "$#,##0.00") & ")"
                    End If

                    If System.Math.Round(AMT_TO_APPLY_total, 2) <> System.Math.Round(PYMT_REF_AMT, 2) Then
                        EMsg &= vbCr & "Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ")" & vbCrLf & " does not equal the Total Chargebacks applied (" & Format(AMT_TO_APPLY_total, "$#,##0.00") & ")"
                        ' AK SAYS SHE WILL NEVER USE THIS SCREEN WITHOUT A CHARGEBACK, AND SHE WILL NEVER WANT TO LEAVE A BALANCE OPEN ON THE CREDIT GENERATED - SHE WILL LEAVE THE BALANCE OPEN ON THE CHARGEBACKS INDIVIDUALLY

                        ' note that the payments tab is based on ARTPYMT1/2/3 - so a ASP payment without application to chargebacks won't even show up on that screen
                        ' so that screen probably needs to be re-architected to show records from SPTAPMT1

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
                    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTACOMC").Select("APPLY = '1'"), "OPS_YYYYPP").Select
                        Dim OPS_YYYYPP As String = row.Item("OPS_YYYYPP")
                        If dst.Tables("SPTACOMC").Select("OPS_YYYYPP = '" & OPS_YYYYPP & "' and APPLY = '0'").Length <> 0 Then
                            If MsgBox("Some Accruals in " & OPS_YYYYPP & " have been queued for Payment leaving others behind" _
                                      & vbCrLf & vbCrLf & "OK to Continue?", _
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    Next
                End If

                If EMsg = "" Then
                    Dim msg As String = ""

                    msg &= vbCrLf & "- Generate a Credit Memo for " & Format(CREDITED, "$#,##0.00")
                    Dim ASP_COMM_OFFSET As Decimal = Val(numASP_COMM_OFFSET.Value & "")
                    msg &= vbCrLf & "- Offset ASP Commission Accruals by " & Format(ASP_COMM_OFFSET, "$#,##0.00")
                    Dim ASP_COMM_ADD_EXP As Decimal = Val(numASP_COMM_ADD_EXP.Value & "")
                    msg &= vbCrLf & "- Record " & IIf(ASP_COMM_ADD_EXP > 0, "additional", "a reduction to") & " ASP Commission Expenses of " & Format(ASP_COMM_ADD_EXP, "$#,##0.00")
                    Dim DEDUCTED As Decimal = Val(numDEDUCTED.Value & "")
                    msg &= vbCrLf & "- Apply " & Format(DEDUCTED, "$#,##0.00") & " of the Credit Memo Generated against Open ASP Chargebacks"
                    msg &= vbCrLf & vbCrLf & "leaving the Credit Memo Generated with an Open Balance of " & Format(CREDITED - DEDUCTED, "$#,##0.00")

                    Dim BALANCE As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(BALANCE)", "APPLY='1' AND BALANCE<>0") & "")
                    If BALANCE <> 0 Then
                        msg &= vbCrLf & vbCrLf & " and leave payment variances of " & Format(BALANCE, "$#,##0.00") & " as open accruals"
                    End If

                    If MsgBox("You are about to:" & vbCrLf & msg & vbCrLf & vbCrLf & "OK to Proceed?", _
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Reverse"
                If rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Then
                    EMsg &= vbCr & "This Payment was Reversed by Payment No " & rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO")
                End If
                If rowSPTAPMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                    EMsg &= vbCr & "This Payment was used to Reverse Payment No " & rowSPTAPMT1.Item("REVERSED_PYMT_NO")
                End If

                If EMsg = "" Then

                    If MsgBox("You are about to Reverse this Payment" & vbCrLf & "OK to Proceed?", _
                            MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reverse"
                Reverse_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
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
                    .Items("Enter Payment").Settings.Enabled = IIf(ScreenMode And Not (EntryMode = "N"), _
                                                         DefaultableBoolean.True, _
                                                         DefaultableBoolean.False)

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Enter Payment").Visible = ScreenMode And (EntryMode = "L")
                    .Items("Reverse").Visible = ScreenMode And (EntryMode = "V")

                    .Items("Done").Visible = Not (EntryMode = "N")
                    .Items("Update").Visible = ScreenMode And (EntryMode = "N")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N")
                End With

                .Groups("Totals").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")
                .Groups("Customer Document").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "V")
                .Groups("Chargebacks").Visible = False

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(grpCustomerDocument, (EntryMode <> "N"))

        grpStores.Visible = Not tf
        grpARTCUST1.Visible = tf
        Set_Read_Only(grpARTCUST1, True)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdSPTACOMC, grdSPTACOML, grdSPTACOMP, grdSPTACOMS, grdSPTACOMX, grdARTOPENX, grdSPTAPMTX}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            If grd.Name = "grdSPTACOMC" Or grd.Name = "grdARTOPENX" Then
                If EntryMode = "N" Then
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

        With grdSPTACOMC.DisplayLayout.Bands(0)
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
        lblReversal.Text = ""

        If ScreenMode Then
            If EntryMode = "V" Then
                grdARTOPENX.Text = "Chargebacks Applied to this ASP Payment"
                grdSPTACOMC.Text = "ASP Commission Accruals used in this ASP Payment"
                tabDetails.Tabs("Open Chargebacks").Text = "Chargebacks Applied"
                tabDetails.Tabs("Paid Previously").Visible = False

                If rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Then
                    lblReversal.Text = "This Payment was Reversed by Payment No " & rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO")
                End If
                If rowSPTAPMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                    lblReversal.Text = "This Payment was used to Reverse Payment No " & rowSPTAPMT1.Item("REVERSED_PYMT_NO")
                End If

            Else
                grdARTOPENX.Text = "Customer Deductions Charged Back"
                grdSPTACOMC.Text = "Open ASP Commission Accruals to Apply"
                tabDetails.Tabs("Open Chargebacks").Text = "Open Chargebacks"
                tabDetails.Tabs("Paid Previously").Visible = True
            End If

            If EntryMode = "N" Then
                If rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Then
                    lblStatus.Text = "Reversed by " & rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO")
                ElseIf rowSPTAPMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                    lblStatus.Text = "Reversing " & rowSPTAPMT1.Item("REVERSED_PYMT_NO")
                Else
                    '  lblStatus.Text = "Issued " & Format(rowSPTAPMT1.Item("INIT_DATE"), "MM/dd/yy") & " by " & rowSPTAPMT1.Item("INIT_OPER")
                    lblStatus.Text = "In Process"
                End If
                lblStatus.Visible = True
            Else
                lblStatus.Visible = False
            End If

            lblPYMT_CTL_NO.Visible = (EntryMode <> "N")

            Setup_tabDetails()
            Set_CB()
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTACOMC", "SPTACOMP", "ARTOPENX", "SPTAPMT1", "SPTAPMT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_Open_Accruals()
        Absx1.txtFor("CUST_CODE").Text = ""
        'Absx1.txtFor("PYMT_REF_NO").Text = ""
        'Absx1.dteFor("PYMT_REF_DATE").Value = DBNull.Value
        'Absx1.numFor("PYMT_REF_AMT").Value = 0

        ACC_CTL_NOs.Clear()

        tabDetails.Tabs("KOHLS Results").Visible = False
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = HFs("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)

        CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        If EntryMode = "N" Then
            PYMT_NO = ASCMAIN1.Next_Control_No("SPTAPMT1.PYMT_NO")
            rowSPTAPMT1 = dst.Tables("SPTAPMT1").NewRow
            rowSPTAPMT1.Item("PYMT_NO") = PYMT_NO
            rowSPTAPMT1.Item("CUST_CODE") = CUST_CODE
            rowSPTAPMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSPTAPMT1.Item("INIT_DATE") = DATETIME_STAMP
            rowSPTAPMT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("SPTAPMT1").Rows.Add(rowSPTAPMT1)
        ElseIf EntryMode = "V" Then
            rowSPTAPMT1 = Fill_Record("SPTAPMT1", PYMT_NO)
        End If

        EnforceConstraints(False)

        If EntryMode = "N" Or EntryMode = "L" Then
            Fill_Records("SPTACOMC", CUST_CODE)
            Sort_grdColumns(grdSPTACOMC, "ASP_CODE,OPS_YYYYPP")
            Fill_Records("SPTACOMP", CUST_CODE)
        ElseIf EntryMode = "V" Then
            ASCMAIN1.sql = "Select * from SPTACOMC where PYMT_NO = '" & PYMT_NO & "'"
            Fill_Records("SPTACOMC", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSPTACOMC, "ASP_CODE,OPS_YYYYPP")

            For Each row As DataRow In dst.Tables("SPTACOMC").Select("")
                ' PROBABLY SHOULD HAVE ACC_CTL_NO_OPEN REFERENCE IN SPTACOMC TO INDICATE THAT BALANCE WAS LEFT OPEN
                Dim ACC_CTL_NO As String = row.Item("ACC_CTL_NO")
                If ASCDATA1.GetDataRow("Select * from SPTACOMC where ACC_CTL_NO_ORIG = '" & ACC_CTL_NO & "'") IsNot Nothing Then
                    row.Item("LEAVE_OPEN") = "1"
                End If
            Next

            Fill_Records("SPTACOMP", "", True, ASCMAIN1.sql)
        End If

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_ASP_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        If EntryMode = "N" Or EntryMode = "L" Then
            If ASCMAIN1.CLIENT = "AHA" Then
                Fill_Records("ARTOPENX", New String() {CUST_BILL_TO_CUST})
            Else
                Fill_Records("ARTOPENX", New String() {REASON_CODE, CUST_BILL_TO_CUST})
            End If
        ElseIf EntryMode = "V" Then
            PYMT_BATCH_NO = rowSPTAPMT1.Item("PYMT_BATCH_NO")
            ASCMAIN1.sql = sqlARTOPENX & " from ARTOPEN1 where CUST_CODE = '" & CUST_CODE & "' and (INV_TYPE,INV_NUM) in (Select INV_TYPE,INV_NUM from ARTPYMT3 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "' and PYMT_BATCH_ILNO <> 0)"
            Fill_Records("ARTOPENX", "", True, ASCMAIN1.sql)
            Fill_Records("ARTOPENX", "", False, Replace(ASCMAIN1.sql, "from ARTOPEN1", "from ARTOPENX ARTOPEN1"))

            ASCMAIN1.sql = "Select * from ARTPYMT3 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("PYMT_BATCH_ILNO <> 0")
                Dim INV_TYPE As String = row.Item("INV_TYPE")
                Dim INV_NUM As String = row.Item("INV_NUM")
                Dim rowARTOPENX As DataRow = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_CODE, INV_TYPE, INV_NUM})
                rowARTOPENX.Item("INV_BALANCE") = row.Item("INV_BALANCE")
                rowARTOPENX.Item("AMT_TO_APPLY") = row.Item("INV_PMT")
            Next

            For Each row As DataRow In dst.Tables("SPTACOMC").Select("")
                row.Item("APPLY") = "1"
            Next

        End If

        EnforceConstraints(True)

        Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables("SPTACOMB").Rows.Clear()
        dst.Tables("SPTACOMC").AcceptChanges()

        Dim OPS_YYYYPP As String = ""

        Dim ASP_COMM_OFFSET As Decimal = 0
        Dim ASP_COMM_CREDITED As Decimal = 0

        For Each rowSPTACOMC As DataRow In dst.Tables("SPTACOMC").Select("LEAVE_OPEN = '1' AND BALANCE <> 0")
            Dim ACC_CTL_NO_orig As String = rowSPTACOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO As String = Replicate_Accrual(ACC_CTL_NO_orig, True)
        Next

        For Each rowSPTACOMC As DataRow In dst.Tables("SPTACOMC").Select("APPLY = '1'", "CUST_CODE,OPS_YYYYPP")

            Dim ACC_CTL_NO As String = rowSPTACOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO_ACCRUAL As String = rowSPTACOMC.Item("ACC_CTL_NO_ACCRUAL")
            If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                rowSPTACOMC.SetAdded()
            End If

            rowSPTACOMC.Item("OPS_YYYYPP_PAID") = ASCMAIN1.CYP
            rowSPTACOMC.Item("PYMT_NO") = PYMT_NO
            rowSPTACOMC.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSPTACOMC.Item("LAST_DATE") = DATETIME_STAMP

            If OPS_YYYYPP <> rowSPTACOMC.Item("OPS_YYYYPP") Then
                OPS_YYYYPP = rowSPTACOMC.Item("OPS_YYYYPP")
                Fill_Records("SPTACOMB", New String() {CUST_CODE, OPS_YYYYPP}, False)
            End If

            Dim HC_CODE As String = rowSPTACOMC.Item("HC_CODE") & ""
            Dim ASP_CODE As String = rowSPTACOMC.Item("ASP_CODE")
            Dim BRAND_CODE As String = rowSPTACOMC.Item("BRAND_CODE")
            Dim INV_NO As String = rowSPTACOMC.Item("INV_NO")

            Dim AMT_COMM_total As Decimal = Val(rowSPTACOMC.Item("AMT_COMM") & "")
            Dim AMT_COMM_PAID_total As Decimal = Val(rowSPTACOMC.Item("AMT_COMM_PAID") & "")

            ASP_COMM_CREDITED += AMT_COMM_PAID_total

            If rowSPTACOMC.Item("LEAVE_OPEN") & "" = "1" Then
                ASP_COMM_OFFSET += AMT_COMM_PAID_total
                rowSPTACOMC.Item("AMT_COMM_ADJ") = 0
                rowSPTACOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_PAID_total
            Else
                ASP_COMM_OFFSET += AMT_COMM_total
                rowSPTACOMC.Item("AMT_COMM_ADJ") = AMT_COMM_PAID_total - AMT_COMM_total
                rowSPTACOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_total
            End If

            Dim AMT_COMM_PAID_spread As Decimal = 0

            Dim sqlw As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and CUST_CODE = '" & CUST_CODE & "' and ASP_CODE = '" & ASP_CODE & "' and INV_NO = '" & INV_NO & "' and BRAND_CODE = '" & BRAND_CODE & "'" & IIf(HC_CODE = "", "", " and HC_CODE = '" & HC_CODE & "'")
            sqlw &= " and ACC_CTL_NO_ACCRUAL = '" & ACC_CTL_NO_ACCRUAL & "'"
            Dim AMT_COMM_all As Decimal = Val(dst.Tables("SPTACOMB").Compute("SUM(AMT_COMM)", sqlw) & "")

            Dim rows() As DataRow = dst.Tables("SPTACOMB").Select(sqlw, "AMT_COMM")
            Dim r As Integer = 0

            Dim COLLECTION_CODEs As New Dictionary(Of String, Decimal)

            Dim TOTAL_ADJ As Decimal = 0

            For Each rowSPTACOMB As DataRow In rows

                Dim AMT_COMM As Decimal = Val(rowSPTACOMB.Item("AMT_COMM") & "")
                Dim AMT_COMM_PAID As Decimal = Val(rowSPTACOMB.Item("AMT_COMM_PAID") & "")

                Dim F As Decimal = 1
                If AMT_COMM_all <> 0 Then F = AMT_COMM / AMT_COMM_all

                Dim AMT_COMM_PAID_now As Decimal = System.Math.Round(AMT_COMM_PAID_total * F, 2)
                AMT_COMM_PAID_spread += AMT_COMM_PAID_now
                AMT_COMM_PAID += AMT_COMM_PAID_now

                r += 1
                If r = rows.Length And AMT_COMM_PAID_total <> AMT_COMM_PAID_spread Then
                    AMT_COMM_PAID += AMT_COMM_PAID_total - AMT_COMM_PAID_spread
                End If

                rowSPTACOMB.Item("AMT_COMM_PAID") = AMT_COMM_PAID

                If AMT_COMM_PAID_total <> AMT_COMM_total Then
                    Dim COLLECTION_CODE As String = rowSPTACOMB.Item("COLLECTION_CODE")
                    If Not COLLECTION_CODEs.ContainsKey(COLLECTION_CODE) Then
                        COLLECTION_CODEs.Add(COLLECTION_CODE, 0)
                    End If
                    Dim AMT As Decimal = (AMT_COMM_PAID_total - AMT_COMM_total) * F
                    If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                        COLLECTION_CODEs(COLLECTION_CODE) += AMT
                    Else
                        COLLECTION_CODEs(COLLECTION_CODE) += AMT
                    End If
                End If

                'If AMT_COMM_PAID <> AMT_COMM Then
                '    Dim COLLECTION_CODE As String = rowSPTACOMB.Item("COLLECTION_CODE")
                '    If Not COLLECTION_CODEs.ContainsKey(COLLECTION_CODE) Then
                '        COLLECTION_CODEs.Add(COLLECTION_CODE, 0)
                '    End If
                '    If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                '        COLLECTION_CODEs(COLLECTION_CODE) += AMT_COMM_PAID_now
                '    Else
                '        COLLECTION_CODEs(COLLECTION_CODE) += AMT_COMM_PAID - AMT_COMM
                '    End If
                'End If

            Next

            ' Record impact to Commission Adjustment (expense)

            If COLLECTION_CODEs.Count <> 0 Then
                For Each COLLECTION_CODE As String In COLLECTION_CODEs.Keys
                    If COLLECTION_CODEs(COLLECTION_CODE) <> 0 Then
                        Dim rowSPTAPMT2 As DataRow = dst.Tables("SPTAPMT2").NewRow
                        rowSPTAPMT2.Item("PYMT_NO") = PYMT_NO
                        rowSPTAPMT2.Item("ACC_CTL_NO") = rowSPTACOMC.Item("ACC_CTL_NO")
                        rowSPTAPMT2.Item("COLLECTION_CODE") = COLLECTION_CODE
                        If rowSPTACOMC.Item("LEAVE_OPEN") & "" = "1" Then
                            rowSPTAPMT2.Item("AMT_COMM_ADJ") = 0
                        Else
                            rowSPTAPMT2.Item("AMT_COMM_ADJ") = COLLECTION_CODEs(COLLECTION_CODE)
                            TOTAL_ADJ += COLLECTION_CODEs(COLLECTION_CODE)
                        End If
                        dst.Tables("SPTAPMT2").Rows.Add(rowSPTAPMT2)
                    End If
                Next

                If TOTAL_ADJ <> 0 Then
                    If System.Math.Abs(TOTAL_ADJ + (AMT_COMM_total - AMT_COMM_PAID_total)) > 0.1 Then
                        Dim msg As String = "Problem with Spreading More or Less Expense on Accrual " & ACC_CTL_NO & "(" & CStr(TOTAL_ADJ) & ":" & CStr(AMT_COMM_total) & ":" & CStr(AMT_COMM_PAID_total) & ")"
                        MsgBox(msg, vbOKOnly, "Please Call ABS before clicking OK")
                        Throw New Exception(msg)
                    End If

                End If
            End If
        Next


        ' PRIOR TO CLOSING JAN'18, I HAD TO CREATE RECORDS IN SPTAPMT2 IN ORDER FOR THE JOURNAL TO BALANCE
        ' PROBLEMS WERE PYMT 000289 (WHICH I CREATED A SINGLE 0 STUB RECORD FOR IN SPTAPMT2)
        ' AND PYMT 000281 (WHICH i NEEDED TO CREATE RECORDS FOR, SINCE IT HAD AN EXPENSE COMPONENT)
        ' WITHOUT REPRESENTATION IN SPTAPMT2, THESE PAYMENTS DID NOT COME OUT ON THE PYMT JOUNRAL
        ' SO WE MAY NEED TO DO SOMETHING IF THERE IS NO ADJUSTMENT, JUST TO GET A 0 RECORD OUT THERE
        ' THE BELOW SQL HINGES ON A SINGLE SPTACOMC RECORD AND SPREADS THE ADJUSTMENT USING SPTACOMB
        ' I AM NOT SURE OF THE PART WHERE WE DIVIDE BY AMT_COMM: / SPTACOMC.AMT_COMM
        ' BECAUSE i THINK THAT THE AMT COMM MAY BE THE TOTAL COMM ON A 2NDARY RECORD, AND THE AMTS IN SPTACOMB MAY BE ORIGINAL COMM AMTS
        ' AND WHAT IF SPTACOMC.AMT_COMM IS ZERO?


        'INSERT INTO AHA.SPTAPMT2@AHA
        'SELECT SPTACOMC.PYMT_NO, SPTACOMC.ACC_CTL_NO, COLLECTION_CODE, ROUND (X.AMT_COMM * SPTACOMC.AMT_COMM_ADJ / SPTACOMC.AMT_COMM,2)
        'from AHA.SPTACOMC, (
        'SELECT '000281' PYMT_NO, '001142' ACC_CTL_NO, COLLECTION_CODE, SUM (AMT_COMM) AMT_COMM
        ' FROM AHA.SPTACOMB
        ' WHERE (CUST_CODE, BRAND_CODE, OPS_YYYYPP, ASP_CODE) in
        '(Select CUST_CODE, BRAND_CODE, OPS_YYYYPP, ASP_CODE from AHA.SPTACOMC WHERE PYMT_NO = '000281')
        'GROUP BY COLLECTION_CODE
        ') X  where SPTACOMC.PYMT_NO = '000281' and X.PYMT_NO = SPTACOMC.PYMT_NO

        '18 Rows Updated

        'INSERT INTO AHA.SPTAPMT2@AHA VALUES ('000289','001145','AVVAL',0)

        '1 Rows Updated

        'COMMIT



        Record_AR(ASP_COMM_OFFSET, ASP_COMM_CREDITED)


        Update_Record_TDA("SPTACOMC")
        Update_Record_TDA("SPTACOMB")

        Update_Record_TDA("SPTAPMT1")
        Update_Record_TDA("SPTAPMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTOPEN1")

        CommitTrans("Update Complete")

    End Sub

    Sub Reverse_Record()

        BeginTrans()

        dst.Tables("SPTACOMB").Rows.Clear()
        dst.Tables("SPTACOMC").AcceptChanges()

        Dim OPS_YYYYPP As String = ""

        Dim ASP_COMM_OFFSET As Decimal = 0
        Dim ASP_COMM_CREDITED As Decimal = 0

        Dim rowSPTAPMT1_orig As DataRow = dst.Tables("SPTAPMT1").Rows.Find(PYMT_NO)
        Dim PYMT_NO_ORIG As String = PYMT_NO
        PYMT_NO = ASCMAIN1.Next_Control_No("SPTAPMT1.PYMT_NO")

        rowSPTAPMT1.Item("REVERSED_BY_PYMT_NO") = PYMT_NO

        rowSPTAPMT1 = dst.Tables("SPTAPMT1").NewRow
        rowSPTAPMT1.Item("PYMT_NO") = PYMT_NO
        rowSPTAPMT1.Item("CUST_CODE") = CUST_CODE
        rowSPTAPMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowSPTAPMT1.Item("INIT_DATE") = DATETIME_STAMP
        rowSPTAPMT1.Item("INIT_OPER") = ASCMAIN1.USER_ID

        rowSPTAPMT1.Item("REVERSED_PYMT_NO") = PYMT_NO_ORIG
        rowSPTAPMT1.Item("PYMT_REF_NO") = rowSPTAPMT1_orig.Item("PYMT_REF_NO")
        rowSPTAPMT1.Item("PYMT_REF_DATE") = rowSPTAPMT1_orig.Item("PYMT_REF_DATE")
        rowSPTAPMT1.Item("PYMT_REF_AMT") = -1 * Val(rowSPTAPMT1_orig.Item("PYMT_REF_AMT"))

        dst.Tables("SPTAPMT1").Rows.Add(rowSPTAPMT1)

        For Each rowSPTACOMC As DataRow In dst.Tables("SPTACOMC").Select("PYMT_NO = '" & PYMT_NO_ORIG & "'")
            Dim ACC_CTL_NO_ACCRUAL As String = rowSPTACOMC.Item("ACC_CTL_NO_ACCRUAL")
            Dim ACC_CTL_NO As String = rowSPTACOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO_orig As String = ACC_CTL_NO

            Dim rowSPTACOMC_new As DataRow = Nothing

            ASCMAIN1.sql = "Select * from SPTACOMC where ACC_CTL_NO = (Select Max(ACC_CTL_NO) from SPTACOMC where ACC_CTL_NO_ACCRUAL = :PARM1)"
            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ACC_CTL_NO_ACCRUAL})
            If row.Item("ACC_CTL_NO") <> ACC_CTL_NO Then
                rowSPTACOMC_new = dst.Tables("SPTACOMC").Rows.Add(row.ItemArray)
                rowSPTACOMC_new.AcceptChanges()
                ACC_CTL_NO = row.Item("ACC_CTL_NO")
            Else
                ACC_CTL_NO = Replicate_Accrual(ACC_CTL_NO_orig, True)
                rowSPTACOMC_new = dst.Tables("SPTACOMC").Rows.Find(ACC_CTL_NO)
            End If

            rowSPTACOMC_new.Item("APPLY") = "1"
            rowSPTACOMC_new.Item("LEAVE_OPEN") = "1"
            rowSPTACOMC_new.Item("PYMT_NO") = PYMT_NO
            rowSPTACOMC_new.Item("AMT_COMM_PAID") = -1 * Val(rowSPTACOMC.Item("AMT_COMM_PAID") & "")
            rowSPTACOMC_new.Item("AMT_COMM_OFFSET") = -1 * Val(rowSPTACOMC.Item("AMT_COMM_OFFSET") & "")

        Next



        For Each rowSPTACOMC As DataRow In dst.Tables("SPTACOMC").Select("LEAVE_OPEN = '1' AND BALANCE <> 0 AND PYMT_NO = '" & PYMT_NO & "'")
            Dim ACC_CTL_NO_orig As String = rowSPTACOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO As String = Replicate_Accrual(ACC_CTL_NO_orig, True)
            Dim row As DataRow = dst.Tables("SPTACOMC").Rows.Find(ACC_CTL_NO)
            row.Item("OPS_YYYYPP_PAID") = DBNull.Value
            row.Item("PYMT_NO") = DBNull.Value
            row.Item("LAST_OPER") = DBNull.Value
            row.Item("LAST_DATE") = DBNull.Value
        Next
 

        For Each rowSPTACOMC As DataRow In dst.Tables("SPTACOMC").Select("PYMT_NO = '" & PYMT_NO & "'", "CUST_CODE,OPS_YYYYPP")

            Dim ACC_CTL_NO As String = rowSPTACOMC.Item("ACC_CTL_NO")
            Dim ACC_CTL_NO_ORIG As String = rowSPTACOMC.Item("ACC_CTL_NO_ORIG")
            Dim ACC_CTL_NO_ACCRUAL As String = rowSPTACOMC.Item("ACC_CTL_NO_ACCRUAL")
            If ACC_CTL_NOs.Contains(ACC_CTL_NO) Then
                rowSPTACOMC.AcceptChanges()
                rowSPTACOMC.SetAdded()
            End If

            rowSPTACOMC.Item("OPS_YYYYPP_PAID") = ASCMAIN1.CYP
            rowSPTACOMC.Item("PYMT_NO") = PYMT_NO
            rowSPTACOMC.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSPTACOMC.Item("LAST_DATE") = DATETIME_STAMP

            If OPS_YYYYPP <> rowSPTACOMC.Item("OPS_YYYYPP") Then
                OPS_YYYYPP = rowSPTACOMC.Item("OPS_YYYYPP")
                Fill_Records("SPTACOMB", New String() {CUST_CODE, OPS_YYYYPP}, False)
            End If

            Dim HC_CODE As String = rowSPTACOMC.Item("HC_CODE") & ""
            Dim ASP_CODE As String = rowSPTACOMC.Item("ASP_CODE")
            Dim BRAND_CODE As String = rowSPTACOMC.Item("BRAND_CODE")
            Dim INV_NO As String = rowSPTACOMC.Item("INV_NO")

            Dim AMT_COMM_total As Decimal = Val(rowSPTACOMC.Item("AMT_COMM") & "")
            Dim AMT_COMM_PAID_total As Decimal = Val(rowSPTACOMC.Item("AMT_COMM_PAID") & "")

            Dim AMT_COMM_OFFSET_total As Decimal = Val(rowSPTACOMC.Item("AMT_COMM_OFFSET") & "")

            ASP_COMM_CREDITED += AMT_COMM_PAID_total

            ASP_COMM_OFFSET += AMT_COMM_OFFSET_total
            'If rowSPTACOMC.Item("LEAVE_OPEN") & "" = "1" Then
            '    ASP_COMM_OFFSET += AMT_COMM_PAID_total
            '    rowSPTACOMC.Item("AMT_COMM_ADJ") = 0
            '    rowSPTACOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_PAID_total
            'Else
            '    ASP_COMM_OFFSET += AMT_COMM_total
            '    rowSPTACOMC.Item("AMT_COMM_ADJ") = AMT_COMM_PAID_total - AMT_COMM_total
            '    rowSPTACOMC.Item("AMT_COMM_OFFSET") = AMT_COMM_total
            'End If

            Dim AMT_COMM_PAID_spread As Decimal = 0

            Dim sqlw As String = "OPS_YYYYPP = '" & OPS_YYYYPP & "' and CUST_CODE = '" & CUST_CODE & "' and ASP_CODE = '" & ASP_CODE & "' and INV_NO = '" & INV_NO & "' and BRAND_CODE = '" & BRAND_CODE & "'" & IIf(HC_CODE = "", "", " and HC_CODE = '" & HC_CODE & "'")
            sqlw = "ACC_CTL_NO_ACCRUAL = '" & ACC_CTL_NO_ACCRUAL & "'"
            Dim AMT_COMM_all As Decimal = Val(dst.Tables("SPTACOMB").Compute("SUM(AMT_COMM)", sqlw) & "")

            Dim rows() As DataRow = dst.Tables("SPTACOMB").Select(sqlw, "AMT_COMM")
            Dim r As Integer = 0
 
            For Each rowSPTACOMB As DataRow In rows

                Dim AMT_COMM As Decimal = Val(rowSPTACOMB.Item("AMT_COMM") & "")
                Dim AMT_COMM_PAID As Decimal = Val(rowSPTACOMB.Item("AMT_COMM_PAID") & "")

                Dim F As Decimal = 1
                If AMT_COMM_all <> 0 Then F = AMT_COMM / AMT_COMM_all

                Dim AMT_COMM_PAID_now As Decimal = System.Math.Round(AMT_COMM_PAID_total * F, 2)
                AMT_COMM_PAID_spread += AMT_COMM_PAID_now
                AMT_COMM_PAID += AMT_COMM_PAID_now

                r += 1
                If r = rows.Length And AMT_COMM_PAID_total <> AMT_COMM_PAID_spread Then
                    AMT_COMM_PAID += AMT_COMM_PAID_total - AMT_COMM_PAID_spread
                End If

                rowSPTACOMB.Item("AMT_COMM_PAID") = AMT_COMM_PAID
            Next

            ' Record impact to Commission Adjustment (expense)

            ASCMAIN1.sql = "Select * from SPTAPMT2 where PYMT_NO = '" & PYMT_NO_ORIG & "' and ACC_CTL_NO = '" & ACC_CTL_NO_ORIG & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim rowSPTAPMT2 As DataRow = dst.Tables("SPTAPMT2").NewRow
                rowSPTAPMT2.ItemArray = row.ItemArray
                rowSPTAPMT2.Item("PYMT_NO") = PYMT_NO
                rowSPTAPMT2.Item("ACC_CTL_NO") = rowSPTACOMC.Item("ACC_CTL_NO")
                rowSPTAPMT2.Item("AMT_COMM_ADJ") = -1 * Val(rowSPTAPMT2.Item("AMT_COMM_ADJ") & "")
                dst.Tables("SPTAPMT2").Rows.Add(rowSPTAPMT2)
            Next
 
        Next

        Record_AR(ASP_COMM_OFFSET, ASP_COMM_CREDITED, True)

        Update_Record_TDA("SPTACOMC")
        Update_Record_TDA("SPTACOMB")

        Update_Record_TDA("SPTAPMT1")
        Update_Record_TDA("SPTAPMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        Update_Record_TDA("ARTPYMT3")
        Update_Record_TDA("ARTOPEN1")

        CommitTrans("Reversal Complete")

        '   Rollback("")

    End Sub

    Sub Record_AR(ASP_COMM_OFFSET As Decimal,
                  ASP_COMM_CREDITED As Decimal, Optional reverse As Boolean = False)

        Dim PYMT_REF_DATE As Date = Absx1.dteFor("PYMT_REF_DATE").Value
        Dim PYMT_REF_NO As String = Absx1.txtFor("PYMT_REF_NO").Text

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_ASP_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)

        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
        Dim TERM_CODE As String = rowSOTTYPE1.Item("TERM_CODE")
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        Dim PYMT_CTL_NO As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        rowSPTAPMT1.Item("PYMT_CTL_NO") = PYMT_CTL_NO

        Dim R As Decimal = 1 : If reverse Then R = -1

        rowSPTAPMT1.Item("ASP_COMM_OFFSET") = ASP_COMM_OFFSET
        rowSPTAPMT1.Item("ASP_COMM_CREDITED") = ASP_COMM_CREDITED


        Dim ASP_COMM_APPLIED As Decimal = Val(Absx1.numFor("DEDUCTED").Value & "") * R
        Dim ASP_COMM_ADD_EXP As Decimal = Val(Absx1.numFor("ASP_COMM_ADD_EXP").Value & "") * R

        rowSPTAPMT1.Item("ASP_COMM_APPLIED") = ASP_COMM_APPLIED
        rowSPTAPMT1.Item("ASP_COMM_ADD_EXP") = ASP_COMM_ADD_EXP
        rowSPTAPMT1.Item("PYMT_DATE") = PYMT_REF_DATE

        Dim INV_TOTAL_AMOUNT As Decimal = Val(Absx1.numFor("CREDITED").Value & "") * R
        Dim DEDUCTED As Decimal = Val(Absx1.numFor("DEDUCTED").Value & "") * R
        Dim INV_DATE As Date = Now.Date

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

            Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
            rowSPTAPMT1.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO

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
                .Item("PYMT_SOURCE") = "A"
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
                .Item("PYMT_NOTE") = "ASP Payment"
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
                Dim INV_PMT As Decimal = Val(rowARTOPENX.Item("AMT_TO_APPLY") & "") * R
                rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {CUST_CODE, INV_TYPE, INV_NUM}, False, False)
                If rowARTOPEN1 Is Nothing Then
                    Dim SQLW As String = " where CUST_CODE = '" & CUST_CODE & "' and INV_TYPE = '" & INV_TYPE & "' and INV_NUM = '" & INV_NUM & "'"
                    ASCMAIN1.sql = "Insert into ARTOPEN1 Select * from ARTOPENX " & SQLW
                    ASCDATA1.ExecuteSQL()
                    ASCMAIN1.sql = "Update ARTOPEN1 Set OPS_YYYYPP_F = NULL " & SQLW
                    ASCDATA1.ExecuteSQL()
                    ASCMAIN1.sql = "Delete from ARTOPENX " & SQLW
                    rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {CUST_CODE, INV_TYPE, INV_NUM}, False, False)
                End If
                PYMT_BATCH_ILNO += 1
                Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, PYMT_BATCH_NO, PYMT_BATCH_ILNO, INV_PMT, PYMT_REF_NO, PYMT_REF_DATE)
            Next
        End If

        If reverse Then
            ASCMAIN1.Record_Event("SPTAPMT1", PYMT_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "R", "Payment Reversal", "")
        Else
            ASCMAIN1.Record_Event("SPTAPMT1", PYMT_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "P", "Payment Entered", "")
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
                CUST_CODE = Split(key, ":")(0)
                PYMT_NO = Split(key, ":")(1)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("View Payment")

                'Absx1.txtFor("PYMT_NO").Text = key
                'Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SPTAPMT1"
            E.COLUMN_NAME = "PYMT_NO"
            E.CODE_VALUE = Absx1.txtFor("PYMT_NO").Text
            E.DESC_VALUE = "Adv/Comm Payment"
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

        E.TABLE_NAME = "SPTAPMT1"
        E.TABLE_KEY_CAPTION = "Adv & Comm Pymt No"
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
        Load_Popup_Menu(grdSPTACOMX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSPTACOMC, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Apply Selected", "Replicate")
        Load_Popup_Menu(grdSPTACOMP, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Replicate")
        Load_Popup_Menu(grdARTOPENX, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Select All from Same Batch")
        Load_Popup_Menu(grdSPTAPMTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
            Case "grdSPTACOMC"
                tlb_pop.Tools("Apply Selected").SharedProps.Visible = (EntryMode = "N")
                tlb_pop.Tools("Replicate").SharedProps.Visible = (EntryMode = "N")

            Case "grdSPTACOMP"
                tlb_pop.Tools("Replicate").SharedProps.Visible = (EntryMode = "N")

            Case "grdARTOPENX"
                tlb_pop.Tools("Select All from Same Batch").SharedProps.Visible = (EntryMode = "N")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSPTACOMX"
                    'If grdSPTACOMX.Tag = "" Then
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

            Case "Select All from Same Batch"
                If grdARTOPENX.ActiveRow IsNot Nothing Then
                    Dim CUST_CODE As String = grdARTOPENX.ActiveRow.Cells("CUST_CODE").Value & ""
                    Dim INV_TYPE As String = grdARTOPENX.ActiveRow.Cells("INV_TYPE").Value & ""
                    Dim INV_NUM As String = grdARTOPENX.ActiveRow.Cells("INV_NUM").Value & ""

                    Dim rowARTOPENX As DataRow = dst.Tables("ARTOPENX").Rows.Find(New String() {CUST_CODE, INV_TYPE, INV_NUM})
                    If rowARTOPENX IsNot Nothing Then
                        ASCMAIN1.sql = "Select * from ARTPYMT5 where CHARGEBACK_IND = '1' and CHARGEBACK_NO = '" & INV_NUM & "'"
                        Dim rowARTPYMT5 As DataRow = ASCDATA1.GetDataRow()
                        If rowARTPYMT5 IsNot Nothing Then
                            Dim PYMT_BATCH_NO As String = rowARTPYMT5.Item("PYMT_BATCH_NO")
                            Dim PYMT_BATCH_LNO As Int32 = Val(rowARTPYMT5.Item("PYMT_BATCH_LNO"))
                            Dim OUR_REFERENCE As String = rowARTPYMT5.Item("OUR_REFERENCE") & ""

                            ASCMAIN1.Progress("Now Selecting all Chargebacks generated in Batch " & PYMT_BATCH_NO)

                            Dim INV_NUMs As New List(Of String)
                            ASCMAIN1.sql = "Select Distinct CHARGEBACK_NO" & vbCrLf _
                                & " from ARTPYMT5" & vbCrLf _
                                & " where CHARGEBACK_IND = '1'" & vbCrLf _
                                & "   and PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" & vbCrLf _
                                & "   and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
                            If Format(Now, "yyyyMMdd") <= "20160704" Then
                                ASCMAIN1.sql &= " and OUR_REFERENCE = '" & OUR_REFERENCE & "'"
                            End If
                            If CUST_CODE = "KOHLS" Then ' S/B OUR_REFERENCE
                                ASCMAIN1.sql &= " and CUST_REFERENCE in (Select INV_NO from SOTINVH1 where INV_TYPE = 'I' and CUST_CODE = 'KOHLS')"
                            End If
                            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                                Dim CHARGEBACK_NO As String = row.Item("CHARGEBACK_NO") & ""
                                INV_NUMs.Add(CHARGEBACK_NO)
                            Next

                            For Each grow As UltraWinGrid.UltraGridRow In grdARTOPENX.Rows
                                Dim INV_NUM_CB As String = grow.Cells("INV_NUM").Value & ""
                                Dim INV_CUST_PO As String = grow.Cells("INV_CUST_PO").Value & ""

                                If INV_NUMs.Contains(INV_NUM_CB) Then
                                    grow.Cells("AMT_TO_APPLY").Value = grow.Cells("INV_BALANCE").Value
                                    grow.Update()
                                End If
                            Next

                            ASCMAIN1.Progress("")

                        End If
                    End If
                End If
        End Select
    End Sub

#End Region

    Function Replicate_Accrual(ACC_CTL_NO_orig As String, Optional leave_open As Boolean = False) As String
        Dim rowSPTACOMC_orig As DataRow
        If leave_open Then
            rowSPTACOMC_orig = dst.Tables("SPTACOMC").Rows.Find(ACC_CTL_NO_orig)
        Else
            rowSPTACOMC_orig = dst.Tables("SPTACOMP").Rows.Find(ACC_CTL_NO_orig)
        End If
        Dim ACC_CTL_NO As String = ASCMAIN1.Next_Control_No("SPTACOMC.ACC_CTL_NO")
        ACC_CTL_NOs.Add(ACC_CTL_NO)

        Dim rowSPTACOMC As DataRow = dst.Tables("SPTACOMC").NewRow
        rowSPTACOMC.ItemArray = rowSPTACOMC_orig.ItemArray
        With rowSPTACOMC
            .Item("ACC_CTL_NO") = ACC_CTL_NO
            If leave_open Then
                .Item("AMT_COMM") = rowSPTACOMC_orig.Item("BALANCE")
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
        dst.Tables("SPTACOMC").Rows.Add(rowSPTACOMC)

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

#Region "grdSPTACOMX"

    Private Sub grdSPTACOMX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTACOMX.AfterCellUpdate
        'If e.Cell.Column.Key = "CUST_DC_IND" Then
        '    'grdSPTACOMX.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        '    grdSPTACOMX.UpdateData()
        'End If

        'If e.Cell.Value & "" <> "" Then

        'End If

    End Sub

    Private Sub grdSPTACOMX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTACOMX.BeforeRowUpdate
        'If e.Row.IsAddRow Then
        '    e.Row.Cells("CUST_CODE").Value = HFs("CUST_CODE")
        'End If
    End Sub

    Private Sub grdSPTACOMX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTACOMX.ClickCellButton
        'Dim sql_where As String = ""
        'Select Case grdSPTACOMX.ActiveCell.Column.Key
        '    Case "SELL_CODE"
        'End Select

        'grdClickCellButton(grdSPTACOMX, sql_where, False)
    End Sub

    Private Sub grdSPTACOMX_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSPTACOMX.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In grdSPTACOMX.Selected.Rows
        '    If dst.Tables("ARTCUST2").Rows(grow.ListIndex).RowState = DataRowState.Added Then
        '    Else
        '        MsgBox("Cannot Delete Existing Store Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '        Exit For
        '    End If
        'Next
    End Sub

    Private Sub grdSPTACOMX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTACOMX.AfterRowActivate
        'If grdSPTACOMX.ActiveRow.IsAddRow Then
        '    grdSPTACOMX.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        'Else
        '    grdSPTACOMX.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        'End If
    End Sub

    Private Sub grdSPTACOMX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTACOMX.AfterExitEditMode

        '    With grdSPTACOMX
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
        Fill_Records("SPTACOMX")
        Sort_grdColumns(grdSPTACOMX, "OPS_YYYYPP".ToLower)

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_ASP_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")
        If ASCMAIN1.CLIENT = "AHA" Then
            Fill_Records("ARTOPENX", New String() {""})
        Else
            Fill_Records("ARTOPENX", New String() {REASON_CODE, ""})
        End If
        Sort_grdColumns(grdARTOPENX, "INV_NUM")

        Fill_Records("SPTAPMTX")
        Sort_grdColumns(grdSPTAPMTX, "PYMT_NO".ToLower)

    End Sub

    Private Sub grdSPTACOMX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTACOMX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Click_Command("View")
        End If
    End Sub

    Sub Setup_grdSPTACOMC()
        If grdSPTACOMC.ActiveRow Is Nothing OrElse Not grdSPTACOMC.ActiveRow.IsDataRow Then
            grdSPTACOMS.Visible = False
            grdSPTACOML.Visible = False
        Else
            Dim OPS_YYYYPP As String = grdSPTACOMC.ActiveRow.Cells("OPS_YYYYPP").Value
            Dim CUST_CODE As String = grdSPTACOMC.ActiveRow.Cells("CUST_CODE").Value
            Dim HC_CODE As String = grdSPTACOMC.ActiveRow.Cells("HC_CODE").Value & ""
            Dim BRAND_CODE As String = grdSPTACOMC.ActiveRow.Cells("BRAND_CODE").Value & ""
            Dim ASP_CODE As String = grdSPTACOMC.ActiveRow.Cells("ASP_CODE").Value & ""
            Dim INV_NO As String = grdSPTACOMC.ActiveRow.Cells("INV_NO").Value & ""
            Dim ACC_CTL_NO_ACCRUAL As String = grdSPTACOMC.ActiveRow.Cells("ACC_CTL_NO_ACCRUAL").Value & ""

            grdSPTACOMS.Visible = True
            Fill_Records("SPTACOMS", New String() {OPS_YYYYPP, CUST_CODE, HC_CODE, BRAND_CODE, ASP_CODE, INV_NO, ACC_CTL_NO_ACCRUAL})
            Sort_grdColumns(grdSPTACOMS, "CUST_STORE_NO")
            grdSPTACOMS.Text = "Detail by Store - " & OPS_YYYYPP & ":" + BRAND_CODE & ":" & HC_CODE

            grdSPTACOML.Visible = True
            Fill_Records("SPTACOML", New String() {OPS_YYYYPP, CUST_CODE, HC_CODE, BRAND_CODE, ASP_CODE, INV_NO, ACC_CTL_NO_ACCRUAL})
            Sort_grdColumns(grdSPTACOML, "COLLECTION_CODE")
            grdSPTACOML.Text = "Detail by Collection - " & OPS_YYYYPP & ":" + BRAND_CODE & ":" & HC_CODE
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

#Region "grdSPTACOMC"

    Private Sub grdSPTACOMC_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTACOMC.AfterCellUpdate
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

    Private Sub grdSPTACOMC_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTACOMC.AfterRowActivate
        Setup_grdSPTACOMC()
    End Sub
    Private Sub grdSPTACOMC_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOMC.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSPTACOMC_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSPTACOMC.ClickCell
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

    Private Sub grdSPTACOMC_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTACOMC.ClickCellButton
        'If Val(e.Cell.Row.Cells("AMT_COMM_PAID").Value & "") <> 0 Then
        '    e.Cell.Row.Cells("AMT_COMM_PAID").Value = 0
        'Else
        '    e.Cell.Row.Cells("AMT_COMM_PAID").Value = e.Cell.Row.Cells("AMT_COMM").Value
        'End If
        'e.Cell.Row.Update()
    End Sub

    Private Sub grdSPTACOMC_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTACOMC.InitializeRow
        If Val(e.Row.Cells("AMT_COMM_ADD").Value & "") <> 0 Or Val(e.Row.Cells("AMT_COMM_RED").Value & "") <> 0 Then
            e.Row.Cells("AMT_COMM_PAID").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("AMT_COMM_PAID").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSPTACOMC_MouseClick(sender As Object, e As MouseEventArgs) Handles grdSPTACOMC.MouseClick
        'If grdSPTACOMC.ActiveRow IsNot Nothing AndAlso grdSPTACOMC.ActiveCell IsNot Nothing Then
        '    If grdSPTACOMC.ActiveCell.DataChanged And grdSPTACOMC.ActiveCell.Column.Key = "APPLY" Then
        '        grdSPTACOMC.ActiveRow.Update()
        '        If grdSPTACOMC.ActiveCell.Value = "0" Then
        '            grdSPTACOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = DBNull.Value
        '        Else
        '            grdSPTACOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = grdSPTACOMC.ActiveRow.Cells("AMT_COMM").Value
        '        End If
        '        grdSPTACOMC.ActiveRow.Update()
        '    End If
        'End If
    End Sub

    Private Sub grdSPTACOMC_MouseUp(sender As Object, e As MouseEventArgs) Handles grdSPTACOMC.MouseUp
        If grdSPTACOMC.ActiveRow IsNot Nothing AndAlso grdSPTACOMC.ActiveCell IsNot Nothing Then
            If grdSPTACOMC.ActiveCell.DataChanged And grdSPTACOMC.ActiveCell.Column.Key = "APPLY" Then
                grdSPTACOMC.ActiveRow.Update()
                'If grdSPTACOMC.ActiveCell.Value = "0" Then
                '    grdSPTACOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = DBNull.Value
                'Else
                '    grdSPTACOMC.ActiveRow.Cells("AMT_COMM_PAID").Value = grdSPTACOMC.ActiveRow.Cells("AMT_COMM").Value
                'End If
                'grdSPTACOMC.ActiveRow.Update()
            End If
        End If
    End Sub

#End Region
    Sub Display_Totals()
        Dim CREDITED As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM_PAID)", "") & "")
        Dim DEDUCTED As Decimal = Val(dst.Tables("ARTOPENX").Compute("SUM(AMT_TO_APPLY)", "") & "")
        Dim AMT_COMM As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM)", "APPLY='1'") & "")
        Dim AMT_COMM_ADD As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM_ADD)", "APPLY='1'") & "")
        Dim AMT_COMM_RED As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM_RED)", "APPLY='1'") & "")
        Dim BALANCE As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(BALANCE)", "APPLY='1'") & "")

        Dim ASP_COMM_OFFSET As Decimal = AMT_COMM - BALANCE

        numCREDITED.Value = CREDITED
        numDEDUCTED.Value = DEDUCTED
        numDIFFERENCE.Value = DEDUCTED - CREDITED
        numASP_COMM_OFFSET.Value = ASP_COMM_OFFSET
        numASP_COMM_ADD_EXP.Value = AMT_COMM_ADD - AMT_COMM_RED
    End Sub

    Private Sub grdARTOPENX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPENX.InitializeRow
        e.Row.Cells("PAY").Value = "->"
    End Sub

    Private Sub grdSPTAPMTX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTAPMTX.DoubleClickRow

        If e.Row.IsDataRow Then
            PYMT_NO = e.Row.Cells("PYMT_NO").Value
            If PYMT_NO <> "" Then
                rowSPTAPMT1 = LookUp("SPTAPMT1", PYMT_NO)
                If rowSPTAPMT1 IsNot Nothing Then
                    CUST_CODE = rowSPTAPMT1.Item("CUST_CODE")
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                    Click_Command("View Payment")
                End If
            End If
        End If
    End Sub

    Private Sub tab_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab.SelectedTabChanged

    End Sub

    Private Sub chkShowRCOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowRCOnly.CheckedChanged
        Set_CB()
    End Sub

    Private Sub chkShowOnAccounts_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOnAccounts.CheckedChanged
        Set_CB()
    End Sub

    Sub Set_CB()

        grdARTOPENX.Rows.ColumnFilters.ClearAllFilters()

        Dim sqlw As String = ""

        If EntryMode = "V" Or EntryMode = "L" Then
            ' NO FILTER
        Else
            If Not chkShowOnAccounts.Checked Then
                sqlw &= " and INV_TYPE = 'B'"
            End If

            If chkShowRCOnly.Checked Then
                sqlw &= " and REASON_CODE = '" & REASON_CODE & "'"
            End If
        End If
       
        Dim dvw As DataView = DirectCast(grdARTOPENX.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlw, 6)

    End Sub

    Private Sub tabDetails_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        UltraExplorerBar1.Groups("Chargebacks").Visible = ScreenMode And (tabDetails.SelectedTab.Key = "Open Chargebacks") And EntryMode = "N"
    End Sub

    Private Sub cmdAutoMatch_Click(sender As Object, e As EventArgs) Handles cmdAutoMatch.Click

        If CUST_CODE = "KOHLS" Then

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Auto-Applying KOHLS")

            ASCDATA1.ExecuteSQL("Delete from " & SPTAPMTI)
            dst.Tables(SPTAPMTI).Rows.Clear()
            Dim sqlw As String = "AMT_TO_APPLY <> 0"
            For Each row As DataRow In dst.Tables("ARTOPENX").Select(sqlw)
                Dim INV_NUM As String = row.Item("INV_NUM")
                dst.Tables(SPTAPMTI).Rows.Add(New String() {INV_NUM})
            Next


            Update_Record_TDA(SPTAPMTI)

            Fill_Records("SPTAPMTY")

            For Each row As DataRow In dst.Tables("SPTAPMTY").Select("")
                sqlw = "OPS_YYYYPP = '" & row.Item("ORDR_YYYYPP_UPDATED") & "' and BRAND_CODE = '" & row.Item("BRAND_CODE") & "' and INV_NO = '0000000000'"
                For Each ASP_CODE As String In New String() {"ADVBROAD", "ADVCOOP", "DIF", "TESTERS"}
                    Dim AMT As Decimal = Val(row.Item(ASP_CODE) & "")
                    If AMT <> 0 Then


                        Dim rows() As DataRow = dst.Tables("SPTACOMC").Select(sqlw & " and ASP_CODE = '" & ASP_CODE & "'")
                        If rows.Length <> 0 Then

                            Dim AMT_COMM As Decimal = Val(rows(0).Item("AMT_COMM") & "")
                            Dim AMT_COMM_PAID As Decimal = Val(rows(0).Item("AMT_COMM_PAID") & "")

                            'If System.Math.Abs(AMT_COMM - INV_BALANCE) < 0.05 Then
                            rows(0).Item("APPLY") = "1"
                            rows(0).Item("AMT_COMM_PAID") = Val(rows(0).Item("AMT_COMM_PAID") & "") + AMT
                            If AMT_COMM <> AMT_COMM_PAID + AMT Then
                                rows(0).Item("LEAVE_OPEN") = "1"
                            End If
                            'If AMT_COMM <> INV_BALANCE Then
                            '    grow.Cells("AMT_COMM_PAID").Value = Val(grow.Cells("AMT_COMM_PAID").Value & "") + INV_BALANCE - AMT_COMM
                            'End If
                            'grow.Update()
                            row.Item("PAID_" & ASP_CODE) = Val(row.Item("PAID_" & ASP_CODE) & "") + AMT
                           
                            Display_Totals()
                            'End If
                        End If
                    End If
                Next
            Next

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            tabDetails.Tabs("KOHLS Results").Visible = True
            tabDetails.SelectedTab = tabDetails.Tabs("KOHLS Results")
            Exit Sub
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdSPTACOMC.Rows
            If grow.Cells("APPLY").Value & "" <> "1" Then
                Dim INV_NO As String = grow.Cells("INV_NO").Value
                Dim AMT_COMM As String = grow.Cells("AMT_COMM").Value
                'Dim INV_NO As String = grow.Cells("").Value
                ' If INV_NO = "0000899495" Then Stop
                If INV_NO <> "0000000000" Then
                    Dim SQL As String = "INV_TYPE = 'B' and ISNULL(AMT_TO_APPLY,0) = 0 and (INV_CUST_PO = '" & INV_NO & "' or INV_CUST_PO = '" & Mid(INV_NO, 5, 6) & "')"
                    Dim rows() As DataRow = dst.Tables("ARTOPENX").Select(SQL)
                    If rows.Length > 1 Then
                        SQL &= " AND REASON_CODE = '" & REASON_CODE & "'"
                        rows = dst.Tables("ARTOPENX").Select(SQL)
                    End If
                    If rows.Length = 1 Then
                        Dim INV_BALANCE As Decimal = Val(rows(0).Item("INV_BALANCE") & "")
                        If System.Math.Abs(AMT_COMM - INV_BALANCE) < 0.05 Then
                            grow.Cells("APPLY").Value = "1"
                            If AMT_COMM <> INV_BALANCE Then
                                grow.Cells("AMT_COMM_PAID").Value = Val(grow.Cells("AMT_COMM_PAID").Value & "") + INV_BALANCE - AMT_COMM
                            End If
                            grow.Update()
                            rows(0).Item("AMT_TO_APPLY") = rows(0).Item("INV_BALANCE")
                            Display_Totals()
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grdSPTACOMC_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTACOMC.InitializeLayout

    End Sub
End Class