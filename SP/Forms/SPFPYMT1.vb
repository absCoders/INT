Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class SPFPYMT1
    ' REVISIT MULTITASK LOCKS IN NEW AND IN REVERSE OPTIONS
    ' need to be able to click that agreements were received while making a payment
    ' need to open up a new notes field in payment so that verbal oks can be documented, and appear on the pymt report
    ' Show cb's applied on the payment report
    ' audit trail (like when clicking that POA or cust agreement was received)
    Dim PYMT_NO As String
    Dim PYMT_TYPE As String

    Dim CUST_CODE As String
    Dim VEND_CODE As String
    Dim CUST_BILL_TO_CUST As String

    Dim rowARTCUST1 As DataRow
    Dim rowAPTVEND1 As DataRow
    Dim rowARTCUST1_BT As DataRow

    Dim rowSPTPYMT1 As DataRow
    Dim sqlSPTCOOP1 As String
    Dim sqlSPTPYMT2 As String

    Dim APPR_STATUS_CODE_BackColors As New Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As New Dictionary(Of String, System.Drawing.Color)

    Dim SP_PARM_LIMIT_PCT As Decimal
    Dim SP_PARM_LIMIT_AMT As Decimal
    Dim sqlAPTSUBM1 As String
    Dim INVOICE_FROM_EMAIL As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SPFPYMTI" Then
            InquiryMode = True
        End If

        Get_PARM("SPTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("APTPARM1")

        SP_PARM_LIMIT_PCT = Val(ROWs("SPTPARM1").Item("SP_PARM_LIMIT_PCT") & "")
        SP_PARM_LIMIT_AMT = Val(ROWs("SPTPARM1").Item("SP_PARM_LIMIT_AMT") & "")

        With dst
            ASCMAIN1.sql = "Select SPTPYMT1.*" & vbCrLf _
                & " from SPTPYMT1" & vbCrLf _
                & " where SPTPYMT1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select SPTCOOP1.*, SPTAVEH1.VEHICLE_DESC" & vbCrLf _
                & ", X.COLLECTION_CODE" & vbCrLf _
                & " from SPTAVEH1,SPTCOOP1" & vbCrLf _
                & ", (Select AUTH_NO, Min (COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & " from SPTCOOP3 group by AUTH_NO) X" & vbCrLf _
                & " where X.AUTH_NO (+) = SPTCOOP1.AUTH_NO" & vbCrLf _
                & "   and SPTAVEH1.VEHICLE_CODE (+) = SPTCOOP1.VEHICLE_CODE" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE = 'A'"
            sqlSPTCOOP1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, True)
            .Tables("SPTCOOP1").Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")

            Create_TDA(.Tables.Add, "SPTPYMT1", "*")
            Create_TDA(.Tables.Add, "SPTCOOP3", "*", 1, False)

            ASCMAIN1.sql = "Select SPTPYMT2.*, SPTAVEH1.VEHICLE_DESC" & vbCrLf _
                & ", SPTCOOP1.CUST_CODE, SPTCOOP1.AUTH_DATE, SPTCOOP1.AUTH_REQ_BY, SPTCOOP1.SREP_CODE" & vbCrLf _
                & ", SPTCOOP1.CUST_REF_NUM" & vbCrLf _
                & ", SPTCOOP1.APPR_STATUS_CODE, X.FEATURE_DESC" & vbCrLf _
                & ", SPTCOOP1.BOOKING_NAME, SPTCOOP1.CUST_AGR_RECD, SPTCOOP1.PROOF_ADV_RECD, SPTCOOP1.SAMPLE_RECD" & vbCrLf _
                & ", SPTCOOP1.SEASON_CODE, X.COLLECTION_CODE" & vbCrLf _
                & ", SPTCOOP1.VEHICLE_CODE, SPTCOOP1.VEHICLE_CPM, SPTCOOP1.DATE_START, SPTCOOP1.DATE_END" & vbCrLf _
                & ", SPTCOOP1.QTY, SPTCOOP1.OTHER_COST, SPTCOOP1.NOTES, SPTCOOP1.OPEN_AMT, SPTCOOP1.PAID_AMT" & vbCrLf _
                & ", SPTCOOP1.STATUS_CODE, SPTCOOP1.PYMTS" & vbCrLf _
                & " from SPTPYMT2,SPTAVEH1,SPTCOOP1" & vbCrLf _
                & ", (Select AUTH_NO" & vbCrLf _
                & ", Min (COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", Max (FEATURE_DESC) FEATURE_DESC" & vbCrLf _
                & " from SPTCOOP3 group by AUTH_NO) X" & vbCrLf _
                & " where X.AUTH_NO (+) = SPTCOOP1.AUTH_NO" & vbCrLf _
                & "   and SPTAVEH1.VEHICLE_CODE (+) = SPTCOOP1.VEHICLE_CODE" & vbCrLf _
                & "   and SPTCOOP1.AUTH_NO = SPTPYMT2.AUTH_NO"
            sqlSPTPYMT2 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SPTPYMT2", "**", 0, True)
            .Tables("SPTPYMT2").Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
            '.Tables("SPTPYMT2").Columns.Add("OPEN_AMT_CALC", GetType(System.Decimal), "IIF(ISNULL(CLOSED,'0')='1',0,IIF(ISNULL(OPEN_AMT,0)-ISNULL(PYMT_REF_AMT,0)<0,0,ISNULL(OPEN_AMT,0)-ISNULL(PYMT_REF_AMT,0)))")
            .Tables("SPTPYMT2").Columns.Add("OPEN_AMT_CALC", GetType(System.Decimal), "IIF(ISNULL(CLOSED,'0')='1',0,IIF(ISNULL(TOTAL,0)-ISNULL(PAID_AMT,0)-ISNULL(PYMT_REF_AMT,0)<0,0,ISNULL(TOTAL,0)-ISNULL(PAID_AMT,0)-ISNULL(PYMT_REF_AMT,0)))")

            Create_TDA(.Tables.Add, "SPTPYMT3", "*", 0, True)

            Create_Relation("SPTPYMT2", "SPTPYMT3", "PYMT_NO,PYMT_LNO")

            .Tables("SPTPYMT3").Columns.Add("TOTAL", GetType(System.Decimal), "PARENT(SPTPYMT2_SPTPYMT3).TOTAL")
            .Tables("SPTPYMT3").Columns.Add("PYMT_REF_AMT", GetType(System.Decimal), "PARENT(SPTPYMT2_SPTPYMT3).PYMT_REF_AMT")
            .Tables("SPTPYMT3").Columns.Add("DIST_PCT", GetType(System.Decimal), "IIF(ISNULL(TOTAL,0)=0, 0, 100 * ISNULL(DIST_AMT,0) / TOTAL)")
            .Tables("SPTPYMT3").Columns.Add("DIST_AMT_PYMT_CALC", GetType(System.Decimal), "DIST_PCT * PYMT_REF_AMT / 100")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")

            Create_TDA(.Tables.Add, "ARTPYMT3", "*")
            .Tables("ARTPYMT3").Columns.Add("SELECTED")
            .Tables("ARTPYMT3").Columns("SELECTED").DefaultValue = "0"
            .Tables("ARTPYMT3").Columns.Add("CUST_CODE")

            Create_TDA(.Tables.Add, "APTINVH1", "*")
            Create_TDA(.Tables.Add, "APTINVH2", "*")

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False, "", 1)

            ASCMAIN1.sql = "Select SPTPYMT1.PYMT_NO, SPTPYMT1.CUST_CODE, SPTPYMT1.CUST_CODE CUST_BILL_TO_CUST, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SPTPYMT1.PYMT_CTL_NO, SPTPYMT1.PYMT_REF_NO, SPTPYMT1.PYMT_REF_AMT, SPTPYMT1.PYMT_REF_AMT INV_PMT" _
                & ", SOTINVH1.CUST_STORE_NO, SOTINVH1.SREP_CODE, SOTINVH1.TERM_CODE" & vbCrLf _
                & ", SOTINVH1.INV_DATE PYMT_REF_DATE, SOTINVH1.INV_COMMENT PYMT_REF_COMMENT" & vbCrLf _
                & " from SOTINVH1,SPTPYMT1" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTINVC1", "**", 0, False, "", 1)
            With .Tables("SOTINVC1").Columns
                .Add("AR_PARM_KEY")
            End With

            sqlAPTSUBM1 = "Select APTSUBM1.*, APTINVH1.INV_STATUS from APTSUBM1,APTINVH1 where APTINVH1.VOUCHER_NO (+) = APTSUBM1.VOUCHER_NO" & vbCrLf
            ASCMAIN1.sql = sqlAPTSUBM1 ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "APTSUBM1", "**", 0, True, "")
            '.Tables("APTSUBM1").Columns.Add("SEL")
            '.Tables("APTSUBM1").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = sqlAPTSUBM1 & " AND SUBMIT_STATUS = 'U'"
            Fill_Records("APTSUBM1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SPTPYMT2.PYMT_NO, SPTPYMT2.PYMT_LNO, SPTPYMT2.PYMT_REF_AMT" & vbCrLf _
                & ", SPTCOOP1.BOOKING_NAME, SPTCOOP1.VEHICLE_CODE" & vbCrLf _
                & ", ARTPYMT3.INV_TYPE, ARTPYMT3.INV_NUM, ARTPYMT3.INV_DATE" & vbCrLf _
                & ", ARTPYMT3.CUST_STORE_NO, ARTPYMT3.INV_CUST_PO, ARTPYMT3.INV_BALANCE, ARTPYMT3.INV_PMT" & vbCrLf _
                & ", ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE, ARTPYMT2.CUST_PYMT_AMT" & vbCrLf _
                & " from SPTPYMT2,ARTPYMT3,ARTPYMT2,SPTCOOP1" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTINVC2", "**", 0, False, "", 2)

            Create_TDA(.Tables.Add, "TATTERM1", "*", 0)
            Create_TDA(.Tables.Add, "SOTSREP1", "*", 0)

            With .Tables.Add("SOTINVP0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("ADDRESS0")
                .Columns.Add("ADDRESS1")
                .Columns.Add("ADDRESS2")
                .Columns.Add("ADDRESS3")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With

        End With

        Fill_Records("TATTERM1")
        Fill_Records("SOTSREP1")
        Fill_SOTINVP0()

        Set_Read_Only(grpPaymentInfo, True)

        grdSPTPYMT2.DataSource = dst.Tables("SPTPYMT2")
        grdSPTCOOP1.DataSource = dst.Tables("SPTCOOP1")
        grdSPTCOOPX.DataSource = dst.Tables("SPTCOOPX")
        grdARTPYMT3.DataSource = dst.Tables("ARTPYMT3")
        grdAPTSUBM1.DataSource = dst.Tables("APTSUBM1")

        Show_Filter(grdSPTCOOP1, True)

        Create_Summary(grdSPTCOOPX, "PYMT_NO", "Count")
        Create_Summary(grdSPTCOOPX, New String() {"PYMT_REF_AMT"})

        Create_Summary(grdSPTCOOP1, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOP1, New String() {"QTY", "OTHER_COST", "TOTAL", "OPEN_AMT", "PAID_AMT"})

        Create_Summary(grdSPTPYMT2, "AUTH_PNO", "Count")
        Create_Summary(grdSPTPYMT2, New String() {"PYMT_REF_AMT", "QTY", "OTHER_COST", "TOTAL", "OPEN_AMT", "PAID_AMT"})

        Create_Summary(grdARTPYMT3, "PYMT_BATCH_ILNO", "Count")
        Create_Summary(grdARTPYMT3, New String() {"INV_BALANCE", "INV_PMT", "SELECTED"})

        Create_Summary(grdAPTSUBM1, "SUBMIT_CTL_NO", "Count")

        With grdSPTCOOPX.DisplayLayout.Bands("SPTCOOPX")
            .Columns("PYMT_NO").Header.Fixed = True
        End With

        With grdARTPYMT3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SELECTED" Or gcol.Key = "INV_PMT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSPTCOOP1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"SELECT", "OPEN_AMT", "AUTH_NO", "BOOKING_NAME", "VEHICLE_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSPTPYMT2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
               {"PYMT_REF_AMT", "CLOSED", "AUTH_PNO", "OPEN_AMT", "OPEN_AMT_CALC", "NOTES_PYMT_APPR", "AUTH_NO", "BOOKING_NAME", "VEHICLE_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "PYMT_REF_AMT" Or gcol.Key = "CLOSED" Or gcol.Key = "NOTES_PYMT_APPR" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If dst.Tables("SPTCOOP1").Columns.Contains(gcol.Key) Then
                    gcol.Width = grdSPTCOOP1.DisplayLayout.Bands(0).Columns(gcol.Key).Width
                End If
            Next
            .Columns("OPEN_AMT_CALC").Width = .Columns("OPEN_AMT").Width
        End With


        grdAPTSUBM1.DisplayLayout.UseFixedHeaders = True
        With grdAPTSUBM1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key = "SEL" Then
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If
                If gcol.Key = "EMAIL" Then
                    .Columns("EMAIL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If

            Next
            .Override.AllowUpdate = DefaultableBoolean.True
        End With
        grdAPTSUBM1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single

        Sort_grdColumns(grdAPTSUBM1, "SUBMIT_CTL_NO".ToLower)

        Show_Filter(grdARTPYMT3, True)

        ASCMAIN1.Add_Value_List(grdSPTCOOP1, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTPYMT2, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        'ASCMAIN1.Add_Value_List(grdSPTCOOPX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOP1, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTPYMT2, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")

        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "PYMT_TYPE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTPYMT1' and COLUMN_NAME = 'PYMT_TYPE'")

        Set_Read_Only(grpPaymentInfo, True)

        'Bind_Controls(UltraGroupBox1, "SPTPYMT1")
        'Bind_Controls(grpPaymentInfo, "SPTPYMT1")

        APPR_STATUS_CODE_BackColors.Add("A", Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("P", Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("G", Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("X", Color.Empty)

        APPR_STATUS_CODE_ForeColors.Add("A", Color.Green)
        APPR_STATUS_CODE_ForeColors.Add("P", Color.Purple)
        APPR_STATUS_CODE_ForeColors.Add("G", Color.Blue)
        APPR_STATUS_CODE_ForeColors.Add("X", Color.Red)

        ASCMAIN1.Add_Value_List(grdAPTSUBM1, "SUBMIT_STATUS", Nothing, New String() {":", "U:Pending", "P:Processed", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdAPTSUBM1, "INV_STATUS", Nothing, New String() {":", "O:Open", "P:Paid", "D:Deleted"})


        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 24) & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(24)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                If Absx1.dteFor("PYMT_REF_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Payment Date is Mandatory"
                Else
                    If Format(Absx1.dteFor("PYMT_REF_DATE").Value, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
                        EMsg &= vbCr & "Payment Date may not be in the Future"
                    ElseIf Format(Absx1.dteFor("PYMT_REF_DATE").Value, "yyyyMMdd") < Format(Now.AddYears(-1), "yyyyMMdd") Then
                        EMsg &= vbCr & "Payment Date may not be more than 1 year ago"
                    End If
                End If

                If optPYMT_TYPE.Value & "" = "" Then
                    EMsg &= vbCr & "You must specify a Payment Type (Check or Credit)"
                Else
                    PYMT_TYPE = optPYMT_TYPE.Value
                    If PYMT_TYPE = "R" Then

                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        If CUST_CODE = "" Then
                            EMsg &= vbCr & "You must supply a Valid Customer"
                        Else
                            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                            If IsNothing(rowARTCUST1) Then
                                EMsg &= vbCr & "Customer Entered Is Not Valid"
                            Else
                                If rowARTCUST1.Item("CUST_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Customer Entered Is Not Active"
                                End If
                            End If
                        End If

                        CUST_BILL_TO_CUST = Absx1.txtFor("CUST_BILL_TO_CUST").Text
                        If CUST_BILL_TO_CUST = "" Then
                            Absx1.txtFor("CUST_BILL_TO_CUST").Text = CUST_CODE
                            ' EMsg &= vbCr & "You must supply a Valid Bill-To Customer (or leave blank)"
                        Else
                            rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
                            If IsNothing(rowARTCUST1_BT) Then
                                EMsg &= vbCr & "Bill-To Customer Entered Is Not Valid"
                            Else
                                If rowARTCUST1_BT.Item("CUST_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Bill-To Customer Entered Is Not Active"
                                End If
                            End If
                        End If

                    ElseIf PYMT_TYPE = "P" Then
                        VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                        If VEND_CODE = "" Then
                            EMsg &= vbCr & "You must supply a Valid Vendor"
                        Else
                            rowAPTVEND1 = LookUp("APTVEND1", VEND_CODE)
                            If IsNothing(rowAPTVEND1) Then
                                EMsg &= vbCr & "Vendor Entered Is Not Valid"
                            Else
                                If rowAPTVEND1.Item("VEND_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Vendor Entered Is Not Active"
                                End If
                            End If
                        End If
                    End If
                End If

                If grdAPTSUBM1.Selected.Rows.Count = 1 And EMsg = "" Then
                    '    EMsg &= vbCr & "Submitted Pending Invoice Line must be selected for new Voucher"
                    'Else
                    Dim SUBMIT_CTL_NO As String = grdAPTSUBM1.Selected.Rows(0).Cells("SUBMIT_CTL_NO").Value & ""
                    If Not ASCMAIN1.Logical_Lock("APTSUBM1", SUBMIT_CTL_NO) Then Exit Sub
                End If

                If EMsg = "" Then
                    If PYMT_TYPE = "R" Then
                        ASCMAIN1.Logical_Lock("SPTPYMT1C", CUST_CODE)
                        ASCMAIN1.Logical_Lock("APFPYMT2", CUST_CODE)
                    Else
                        ASCMAIN1.Logical_Lock("SPTPYMT1V", VEND_CODE)
                        ASCMAIN1.Logical_Lock("APFPYMT2", VEND_CODE)
                    End If
                End If

            Case "View"
                PYMT_NO = Absx1.txtFor("PYMT_NO").Text
                If PYMT_NO = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowSPTPYMT1 = LookUp("SPTPYMT1", PYMT_NO)
                    If rowSPTPYMT1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & PYMT_NO & " on File"
                    End If
                End If

            Case "Reverse"
                PYMT_NO = Absx1.txtFor("PYMT_NO").Text
                If PYMT_NO = "" Then
                    EMsg &= vbCr & "You must specify a Document No to Reverse"
                Else
                    rowSPTPYMT1 = LookUp("SPTPYMT1", PYMT_NO)
                    If rowSPTPYMT1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & PYMT_NO & " on File"
                    Else
                        If rowSPTPYMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Or rowSPTPYMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                            EMsg &= vbCr & "Cannot Reverse a Payment that is Already Reversed"
                        End If
                    End If
                End If



                EMsg &= vbCr & "This function is not supported (yet)"

                If EMsg = "" Then
                    Stop ' MULTI-TASK
                End If


            Case "Update"
                Dim DT As Date = Absx1.dteFor("PYMT_REF_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Payment Date is Mandatory"
                Else
                    If Now.Subtract(DT).TotalDays < 0 Then
                        EMsg &= vbCr & "Payment Date cannot be in the Future"
                    Else
                        If Now.Subtract(DT).TotalDays > 545 Then
                            EMsg &= vbCr & "Payment Date is more than 18 months in the past"
                        Else
                            If Now.Subtract(DT).TotalDays > 30 Then
                                If MsgBox("Payment Date is more than 30 days prior to today", MsgBoxStyle.YesNo, "OK to Proceed?") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If

                If Absx1.txtFor("PYMT_REF_NO").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Payment Reference (ie Customer Claim)"
                End If


                Dim PYMT_REF_AMT_total As Decimal = Val(dst.Tables("SPTPYMT2").Compute("SUM(PYMT_REF_AMT)", "") & "")
                Dim PYMT_REF_AMT_entered As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")

                If grdSPTPYMT2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Promo Event Agreements selected to be Paid"
                Else
                    If PYMT_REF_AMT_entered <> PYMT_REF_AMT_total Then
                        EMsg &= vbCr & "Total Pymts applied to Promo Event Commitments (" _
                            & Format(PYMT_REF_AMT_total, "$#,##0.00") & ")" & vbCrLf _
                            & " do not match Payment Amount (" _
                            & Format(PYMT_REF_AMT_entered, "$#,##0.00") & ")"
                    End If
                End If

                If PYMT_TYPE = "P" And PYMT_REF_AMT_total <> 0 Then
                    If Absx1.txtFor("VEND_BUYER_CODE").Text = "" Then
                        EMsg &= vbCr & "Buyer is Mandatory to Approve AP Invoice"
                    Else
                        Dim rowPOTBUYR1 As DataRow = LookUp("POTBUYR1", Absx1.txtFor("VEND_BUYER_CODE").Text)
                        If rowPOTBUYR1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Buyer"
                        Else
                            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", Absx1.txtFor("VEND_BUYER_CODE").Text)
                            If rowASTUSER1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Buyer (User record not found)"
                            Else
                                If rowASTUSER1.Item("USER_STATUS") & "" <> "A" Then
                                    EMsg &= vbCr & "Invalid Value Specified for Buyer (User not Active)"
                                End If
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    For Each row As DataRow In dst.Tables("SPTPYMT2").Select("")
                        Dim PAID_AMT As Decimal = Val(row.Item("PAID_AMT") & "")
                        Dim PYMT_REF_AMT As Decimal = Val(row.Item("PYMT_REF_AMT") & "")
                        Dim TOTAL As Decimal = Val(row.Item("TOTAL") & "")

                        Dim OVER As Decimal = PAID_AMT + PYMT_REF_AMT - TOTAL
                        If OVER > SP_PARM_LIMIT_AMT Or (TOTAL = 0 OrElse 100 * OVER / TOTAL > SP_PARM_LIMIT_PCT) Then
                            If MsgBox("There are payments listed that exceeed Allowable Limits" _
                                      & vbCrLf & vbCrLf & "OK to Continue with Payment?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                            Exit For
                        End If
                    Next


                    If PYMT_TYPE = "R" Then
                        If dst.Tables("ARTPYMT3").Select("SELECTED='1'").Length <> 0 Then
                            Dim INV_PMT_total As Decimal = Val(dst.Tables("ARTPYMT3").Compute("SUM(INV_PMT)", "SELECTED='1'") & "")
                            Dim PYMT_REF_AMT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")

                            If INV_PMT_total <> PYMT_REF_AMT Then
                                EMsg &= vbCr & "Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ")" & vbCrLf & " does not equal the Total Chargebacks applied (" & Format(INV_PMT_total, "$#,##0.00") & ")"
                                ' AK SAYS SHE WILL NEVER USE THIS SCREEN WITHOUT A CHARGEBACK, AND SHE WILL NEVER WANT TO LEAVE A BALANCE OPEN ON THE CREDIT GENERATED - SHE WILL LEAVE THE BALANCE OPEN ON THE CHARGEBACKS INDIVIDUALLY

                                'If MsgBox("Chargebacks have been applied." & vbCrLf & vbCrLf & "However, the Total Payment Amount (" & Format(PYMT_REF_AMT, "$#,##0.00") & ")" & vbCrLf & " does not equal the Total Chargebacks applied (" & Format(INV_PMT, "$#,##0.00") & ")." & vbCrLf _
                                '          & "An AR Item will remain on the Customer Account with a " & IIf(-1 * PYMT_REF_AMT + INV_PMT > 0, "DR", "CR") & " Balance of " & Format(-1 * PYMT_REF_AMT + INV_PMT, "$#,##0.00") & "." _
                                '                 & vbCrLf & vbCrLf & "OK to Continue with Payment?", _
                                '                 MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                '    Exit Sub
                                'End If
                            End If

                            For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("SELECTED='1'")
                                Dim INV_PMT As Decimal = Val(rowARTPYMT3.Item("INV_PMT") & "")
                                Dim INV_BALANCE As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE") & "")
                                Dim INV_BALANCE_NEW As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE_NEW") & "")
                                If (INV_BALANCE >= 0 And (INV_PMT < 0 Or INV_PMT > INV_BALANCE)) _
                                Or (INV_BALANCE <= 0 And (INV_PMT > 0 Or INV_PMT < INV_BALANCE)) Then
                                    EMsg &= vbCr & "Invalid Payment Amount for AR Item " & rowARTPYMT3.Item("INV_TYPE") & "-" & rowARTPYMT3.Item("INV_NUM")
                                End If
                            Next
                        Else
                            EMsg &= vbCr & "You must apply the Payment amount to Chargebacks - why else would you be in this screen?"
                            ' AK SAYS SHE WILL NEVER USE THIS SCREEN WITHOUT A CHARGEBACK, AND SHE WILL NEVER WANT TO LEAVE A BALANCE OPEN ON THE CREDIT GENERATED - SHE WILL LEAVE THE BALANCE OPEN ON THE CHARGEBACKS INDIVIDUALLY

                        End If
                    End If
                End If


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)


                lblSUBMITTED_INVOICE1.Visible = False
                lblSUBMITTED_INVOICE2.Visible = False
                lblSUBMITTED_INVOICE3.Visible = False
                INVOICE_FROM_EMAIL = ""


                If grdAPTSUBM1.Selected.Rows.Count = 1 Then

                    Dim grow As UltraWinGrid.UltraGridRow = grdAPTSUBM1.Selected.Rows(0)

                    lblSUBMITTED_INVOICE1.Text = "From: " & grow.Cells("SUBMIT_EMAIL_FROM").Value
                    lblSUBMITTED_INVOICE1.Visible = True
                    lblSUBMITTED_INVOICE2.Text = "Subj: " & grow.Cells("SUBMIT_SUBJECT").Value
                    lblSUBMITTED_INVOICE2.Visible = True
                    lblSUBMITTED_INVOICE3.Text = "Submitted:  " & grow.Cells("SUBMIT_DATE_RECEIVED").Value
                    lblSUBMITTED_INVOICE3.Visible = True

                    INVOICE_FROM_EMAIL = grow.Cells("SUBMIT_CTL_NO").Value
                    ASCMAIN1.sql = sqlAPTSUBM1 & " and SUBMIT_CTL_NO = '" & INVOICE_FROM_EMAIL & "'" & vbCrLf
                    Fill_Records("APTSUBM1", "", True, ASCMAIN1.sql)
                End If


            Case "Reverse"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Reconciling Invoice"
                Print_Reconciling_Invoice()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode

                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Reverse").Settings.Enabled = DefaultableBoolean.True
                        If rowSPTPYMT1.Item("REVERSED_PYMT_NO") & "" = "" And rowSPTPYMT1.Item("REVERSED_BY_PYMT_NO") & "" = "" Then
                            .Items("Reverse").Visible = True
                        Else
                            .Items("Reverse").Visible = False
                        End If
                    Else
                        ' .Items("Reverse").Settings.Enabled = not_iScreenMode
                        .Items("Reverse").Visible = False
                    End If

                    .Items("Reverse").Visible = False ' not supported yet

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Print").Visible = (EntryMode = "V" And ScreenMode)
                    .Items("Reconciling Invoice").Visible = (EntryMode = "V" And ScreenMode) AndAlso rowSPTPYMT1.Item("PYMT_TYPE") & "" = "R"

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

                End With

                .Groups("Show if Posted in").Visible = Not ScreenMode
                .Groups("Payment Info").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.dteFor("PYMT_REF_DATE"), False)
        If ASCMAIN1.CLIENT = "INT" Then Set_Read_Only_for_ctl(Absx1.txtFor("CUST_BILL_TO_CUST"), True)

        SplitContainer1.Visible = ScreenMode
        grdSPTCOOPX.Visible = Not ScreenMode
        SplitContainer2.Visible = Not ScreenMode

        If ScreenMode Then
            lblVEND_BUYER_CODE.Visible = (PYMT_TYPE = "P")
            txtVEND_BUYER_CODE.Visible = (PYMT_TYPE = "P")

            If rowSPTPYMT1.Item("REVERSED_BY_PYMT_NO") & "" <> "" Then
                lblStatus.Text = "Reversed by " & rowSPTPYMT1.Item("REVERSED_BY_PYMT_NO")
            ElseIf rowSPTPYMT1.Item("REVERSED_PYMT_NO") & "" <> "" Then
                lblStatus.Text = "Reversing " & rowSPTPYMT1.Item("REVERSED_PYMT_NO")
            Else
                lblStatus.Text = "Issued " & Format(rowSPTPYMT1.Item("INIT_DATE"), "MM/dd/yy") & " by " & rowSPTPYMT1.Item("INIT_OPER")
            End If

            lblPYMT_CTL_NO.Visible = (EntryMode <> "N")

            If PYMT_TYPE = "R" Then
                lblPYMT_CTL_NO.Text = "CR Memo"
            Else
                lblPYMT_CTL_NO.Text = "Voucher"
            End If

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpPaymentInfo, (EntryMode = "V"))

            If EntryMode = "N" Then
                With grdARTPYMT3.DisplayLayout.Bands(0)
                    .Columns("SELECTED").Hidden = False
                    .Columns("INV_BALANCE_NEW").Hidden = True
                End With

                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTPYMT2, grdSPTCOOP1, grdARTPYMT3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        If grd.Name = "grdSPTPYMT2" Then
                            .AllowDelete = DefaultableBoolean.True
                        Else
                            .AllowDelete = DefaultableBoolean.False
                        End If
                        If grd.Name = "grdSPTCOOP1" Then
                            .AllowUpdate = DefaultableBoolean.False
                        Else
                            .AllowUpdate = DefaultableBoolean.True
                        End If
                    End With
                Next
                'With grdSPTCOOP1.DisplayLayout.Bands(0)
                '    .Columns("VEHICLE_CODE").CellAppearance.BackColor = Color.LightYellow
                '    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.LightYellow
                'End With

            Else
                With grdARTPYMT3.DisplayLayout.Bands(0)
                    .Columns("SELECTED").Hidden = True
                    .Columns("INV_BALANCE_NEW").Hidden = False
                End With

                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTPYMT2, grdSPTCOOP1, grdARTPYMT3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                'With grdSPTCOOP1.DisplayLayout.Bands(0)
                '    .Columns("VEHICLE_CODE").CellAppearance.BackColor = Color.Empty
                '    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.Empty
                'End With
            End If

            grdSPTCOOP1.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = (PYMT_TYPE = "R")

            With grdSPTPYMT2.DisplayLayout.Bands(0)
                .Columns("CUST_CODE").Hidden = (PYMT_TYPE = "R")
                .Columns("OPEN_AMT").Hidden = (EntryMode = "V")
                .Columns("OPEN_AMT_CALC").Hidden = (EntryMode = "V")
                .Columns("CLOSED").Hidden = (EntryMode = "V")
            End With

            tabDetails.Tabs("Open Commitments").Visible = (EntryMode = "N")
            tabDetails.Tabs("Customer Deductions Charged Back").Visible = (PYMT_TYPE = "R")
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTPYMT1", "SPTPYMT2", "SPTPYMT3",
                                                       "SPTCOOP1", "ARTOPEN1",
                                                       "ARTPYMT1", "ARTPYMT2", "ARTPYMT3",
                                                       "APTINVH1", "APTINVH2", "SOTINVC1", "SOTINVC2", "APTSUBM1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optPYMT_TYPE.Value = "R"

        lblSUBMITTED_INVOICE1.Visible = False
        lblSUBMITTED_INVOICE2.Visible = False
        lblSUBMITTED_INVOICE3.Visible = False
        INVOICE_FROM_EMAIL = ""

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowSPTPYMT1 = dst.Tables("SPTPYMT1").NewRow
            PYMT_NO = ASCMAIN1.Next_Control_No("SPTPYMT1.PYMT_NO")
            With rowSPTPYMT1
                .Item("PYMT_NO") = PYMT_NO
                .Item("PYMT_TYPE") = PYMT_TYPE
                .Item("PYMT_REF_DATE") = HFs("PYMT_REF_DATE")
                If PYMT_TYPE = "R" Then
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
                Else
                    .Item("VEND_CODE") = VEND_CODE
                End If
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With

            dst.Tables("SPTPYMT1").Rows.Add(rowSPTPYMT1)
            ASCMAIN1.sql = sqlSPTCOOP1
            ' NEED TO SEE ALL UNTIL WE ARE FINISHED NETTING
            'ASCMAIN1.sql = sqlSPTCOOP1 _
            '            & " and SPTCOOP1.STATUS_CODE = 'O'" _
            '            & " and NVL(SPTCOOP1.OPEN_AMT,0) <> 0"

            ' MAYBE THIS IS WHERE WE DO APPR_STATUS_CODE = 'A' INSTEAD OF WHEN SETTING UP THE ADAPTOR
            If PYMT_TYPE = "R" Then
                ASCMAIN1.sql &= " and SPTCOOP1.CUST_CODE = '" & CUST_CODE & "'"
            End If
            Fill_Records("SPTCOOP1", "", True, ASCMAIN1.sql)

            dst.Tables("SPTPYMT2").Rows.Clear()

        Else
            rowSPTPYMT1 = Fill_Record("SPTPYMT1", PYMT_NO)
            PYMT_TYPE = rowSPTPYMT1.Item("PYMT_TYPE")

            ASCMAIN1.sql = sqlSPTPYMT2 & " and SPTPYMT2.PYMT_NO = '" & PYMT_NO & "'"
            Fill_Records("SPTPYMT2", "", True, ASCMAIN1.sql)

            dst.Tables("ARTPYMT3").Rows.Clear()

            ASCMAIN1.sql = "Select * from APTSUBM1 where VOUCHER_NO = '" & rowSPTPYMT1.Item("PYMT_CTL_NO") & "'"
            Dim rowAPTSUBM1 As DataRow = ASCDATA1.GetDataRow
            Dim PO_QTY_OPN As Integer = 0
            If rowAPTSUBM1 IsNot Nothing Then
                lblSUBMITTED_INVOICE1.Text = "From: " & rowAPTSUBM1.Item("SUBMIT_EMAIL_FROM") & ""
                lblSUBMITTED_INVOICE1.Visible = True
                lblSUBMITTED_INVOICE2.Text = "Subj: " & rowAPTSUBM1.Item("SUBMIT_SUBJECT") & ""
                lblSUBMITTED_INVOICE2.Visible = True
                lblSUBMITTED_INVOICE3.Text = "Submitted:  " & rowAPTSUBM1.Item("SUBMIT_DATE_RECEIVED") & ""
                lblSUBMITTED_INVOICE3.Visible = True
            End If

        End If

        CUST_CODE = rowSPTPYMT1.Item("CUST_CODE") & ""
        CUST_BILL_TO_CUST = rowSPTPYMT1.Item("CUST_BILL_TO_CUST") & ""
        VEND_CODE = rowSPTPYMT1.Item("VEND_CODE") & ""

        If PYMT_TYPE = "R" Then
            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
            If CUST_BILL_TO_CUST = "" Then
                CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                If CUST_BILL_TO_CUST = "" Then
                    CUST_BILL_TO_CUST = CUST_CODE
                End If
                rowSPTPYMT1.Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            End If

            rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

            If EntryMode = "N" Then
                Load_ChargeBacks()
            Else
                Load_Chargebacks_Applied()
            End If
        Else
            rowAPTVEND1 = LookUp("APTVEND1", VEND_CODE)
            Absx1.txtFor("VEND_BUYER_CODE").Text = rowAPTVEND1.Item("VEND_BUYER_CODE") & ""
        End If

        DisplayTotals()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        If PYMT_TYPE = "R" Then
            ASCDATA1.DeleteRows("ARTPYMT3", "SELECTED <> '1'")
        End If

        For Each rowSPTPYMT3 As DataRow In dst.Tables("SPTPYMT3").Select("")
            rowSPTPYMT3.Item("DIST_AMT_PYMT") = rowSPTPYMT3.Item("DIST_AMT_PYMT_CALC")
        Next

        Dim PYMT_REF_DATE As Date = Absx1.dteFor("PYMT_REF_DATE").Value
        Dim PYMT_REF_NO As String = Absx1.txtFor("PYMT_REF_NO").Text

        Dim ORDR_TYPE_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_PROMO_TYPE_CODE")
        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ORDR_TYPE_CODE)

        '  Dim ChargeBackTotal As Decimal = Val(dst.Tables("ARTPYMT3").Compute("SUM(INV_PMT)", "") & "")

        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE") ' ROWs("SPTPARM1").Item("SP_PARM_AR_POST_CODE") & ""
        Dim TERM_CODE As String = rowSOTTYPE1.Item("TERM_CODE") ' ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0") & ""
        Dim REASON_CODE As String = rowSOTTYPE1.Item("REASON_CODE")

        Dim PYMT_CTL_NO As String = ""
        If PYMT_TYPE = "R" Then
            PYMT_CTL_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        Else
            PYMT_CTL_NO = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        End If

        rowSPTPYMT1.Item("PYMT_CTL_NO") = PYMT_CTL_NO

        Dim VOUCHER_LNO As Integer = 0

        For Each rowSPTPYMT2 As DataRow In dst.Tables("SPTPYMT2").Select("PYMT_REF_AMT <> 0 or CLOSED = '1'")
            With rowSPTPYMT2
                Dim AUTH_NO As String = .Item("AUTH_NO")

                Dim rowSPTCOOP1 As DataRow = dst.Tables("SPTCOOP1").Rows.Find(New Object() {AUTH_NO})
                Dim PYMT_REF_AMT As Decimal = Val(.Item("PYMT_REF_AMT") & "")
                rowSPTCOOP1.Item("PAID_AMT") = Val(rowSPTCOOP1.Item("PAID_AMT") & "") + PYMT_REF_AMT
                Dim OPEN_AMT As Decimal = Val(rowSPTCOOP1.Item("OPEN_AMT") & "") - PYMT_REF_AMT

                If OPEN_AMT <= 0 Or .Item("CLOSED") & "" = "1" Then
                    OPEN_AMT = 0
                    rowSPTCOOP1.Item("STATUS_CODE") = "C"
                End If
                rowSPTCOOP1.Item("OPEN_AMT") = OPEN_AMT
                rowSPTCOOP1.Item("PYMTS") = Val(rowSPTCOOP1.Item("PYMTS") & "") + 1
                'If .Item("CLOSED") & "" = "1" Then
                '    Write_Audit(rowSPTCOOP1)
                'End If

                If PYMT_TYPE = "P" Then
                    Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
                    With rowAPTINVH2
                        .Item("VOUCHER_NO") = PYMT_CTL_NO
                        VOUCHER_LNO = VOUCHER_LNO + 1
                        .Item("VOUCHER_LNO") = VOUCHER_LNO
                        .Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_PROMO_ACCT_CODE_APX")
                        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                        .Item("INV_LINE_AMT") = PYMT_REF_AMT
                    End With
                    dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
                End If
            End With
        Next

        If PYMT_TYPE = "P" Then
            Dim INV_AMT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")

            Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").NewRow
            With rowAPTINVH1
                .Item("VOUCHER_NO") = PYMT_CTL_NO
                .Item("VEND_CODE") = VEND_CODE
                .Item("INV_TYPE") = "I"
                .Item("INV_NUM") = PYMT_REF_NO
                .Item("INV_DATE") = PYMT_REF_DATE

                .Item("INV_AMT") = INV_AMT
                .Item("INV_REF") = PYMT_NO

                .Item("VEND_CODE_AP") = rowAPTVEND1.Item("VEND_CODE_AP")
                If rowAPTVEND1.Item("VEND_PYMT_ADDR") & "" = "" Then
                    .Item("VEND_ALT_CODE") = ""
                    .Item("INV_REMIT_TO") = "V"
                Else
                    .Item("VEND_ALT_CODE") = rowAPTVEND1.Item("VEND_PYMT_ADDR")
                    .Item("INV_REMIT_TO") = "A"
                End If

                .Item("INV_SEP_CHECK") = rowAPTVEND1.Item("VEND_SEP_CHECKS")
                .Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")

                If rowAPTVEND1.Item("BANK_CODE") & "" = "" Then
                    .Item("BANK_CODE") = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
                Else
                    .Item("BANK_CODE") = rowAPTVEND1.Item("BANK_CODE")
                End If

                If rowAPTVEND1.Item("VEND_PYMT_METHOD") & "" = "" Then
                    If .Item("BANK_CODE") & "" <> "" Then
                        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", .Item("BANK_CODE"))
                        .Item("INV_PYMT_METHOD") = rowGLTBANK1.Item("BANK_PYMT_METHOD")
                    End If
                Else
                    .Item("INV_PYMT_METHOD") = rowAPTVEND1.Item("VEND_PYMT_METHOD")
                End If

                .Item("INV_PYMT_CYCLE") = rowAPTVEND1.Item("VEND_PYMT_CYCLE")

                If rowAPTVEND1.Item("POST_CODE") & "" <> "" Then
                    .Item("POST_CODE") = rowAPTVEND1.Item("POST_CODE")
                Else
                    .Item("POST_CODE") = ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
                End If

                .Item("INV_STATUS") = "O"
                .Item("INV_PYMT_CYCLE") = DBNull.Value
                .Item("INV_DUE_DATE") = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, rowAPTVEND1.Item("TERM_CODE") & "", Nothing, .Item("INV_DATE"))
                .Item("INV_BALANCE") = INV_AMT
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1

                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                .Item("REGISTER_IND") = "0"
                .Item("INV_BL_DATE") = .Item("INV_DATE")
                .Item("INV_AMT_VEND") = .Item("INV_AMT")

                If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                    .Item("INV_APPR_STATUS") = "A"
                    Write_Event_Log("APTINVH1", PYMT_CTL_NO, "Auto Approved")
                Else
                    .Item("INV_APPR_STATUS") = "P"
                End If

                .Item("VEND_BUYER_CODE") = Absx1.txtFor("VEND_BUYER_CODE").Text
            End With

            dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)
        End If

        If PYMT_TYPE = "R" Then
            Dim INV_TOTAL_AMOUNT As Decimal = Val(Absx1.numFor("PYMT_REF_AMT").Value & "")

            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
            With rowARTOPEN1
                .Item("CUST_CODE") = CUST_BILL_TO_CUST
                .Item("INV_TYPE") = "C"
                .Item("INV_NUM") = PYMT_CTL_NO
                .Item("INV_DATE") = PYMT_REF_DATE
                .Item("POST_CODE") = POST_CODE
                .Item("TERM_CODE") = TERM_CODE
                .Item("INV_DUE_DATE") = PYMT_REF_DATE
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

            Dim REASON_CODE_CB As String = ""

            If dst.Tables("ARTPYMT3").Select("").Length <> 0 Then

                Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
                rowSPTPYMT1.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO

                Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
                With rowARTPYMT1
                    .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                    .Item("PYMT_BATCH_DATE") = PYMT_REF_DATE
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "1"
                    .Item("PYMT_APPL_ONLY") = "1"
                    .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                    .Item("CURR_EXCH_RATE") = 1
                    .Item("PYMT_SOURCE") = "C"
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
                    .Item("PYMT_NOTE") = "Promo Event Payment"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                End With
                dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

                Dim PYMT_BATCH_ILNO As Integer = 0
                For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("", "PYMT_BATCH_ILNO")
                    PYMT_BATCH_ILNO += 1
                    rowARTPYMT3.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                    rowARTPYMT3.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
                Next

                Dim row As DataRow = Write_ARTPYMT3_from_ARTOPEN1(rowARTOPEN1, 0)
                row.Item("INV_PMT") = -1 * Val(dst.Tables("ARTPYMT3").Compute("SUM(INV_PMT)", "") & "")
                row.Item("INV_PMT_CURR") = row.Item("INV_PMT")
                row.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO

                '   Dim rowARTOPEN1 As DataRow = Nothing

                For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("", "PYMT_BATCH_ILNO")
                    With rowARTPYMT3
                        Dim INV_TYPE As String = .Item("INV_TYPE")
                        Dim INV_NUM As String = .Item("INV_NUM")
                        Dim INV_PMT As Decimal = Val(.Item("INV_PMT") & "")
                        PYMT_BATCH_ILNO = Val(.Item("PYMT_BATCH_ILNO") & "")

                        If REASON_CODE_CB = "" Then
                            REASON_CODE_CB = rowARTPYMT3.Item("REASON_CODE")
                            row.Item("REASON_CODE") = REASON_CODE_CB
                        End If

                        If PYMT_BATCH_ILNO > 0 Then
                            rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() {CUST_BILL_TO_CUST, INV_TYPE, INV_NUM}, False, False)
                        End If

                        Dim INV_BALANCE As Decimal = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
                        rowARTOPEN1.Item("INV_PMT") = Val(rowARTOPEN1.Item("INV_PMT") & "") + INV_PMT
                        rowARTOPEN1.Item("INV_BALANCE") = INV_BALANCE - INV_PMT

                        ' rowARTOPEN1.Item("INV_LAST_PMT") = DATETIME_STAMP.Date
                        rowARTOPEN1.Item("INV_LAST_PMT_REF") = PYMT_REF_NO
                        rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = PYMT_REF_DATE
                        rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowARTOPEN1.Item("LAST_DATE") = DATETIME_STAMP
                        rowARTOPEN1.Item("INV_PMT_CURR") = rowARTOPEN1.Item("INV_PMT")
                        rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")

                        .Item("INV_BALANCE_NEW") = rowARTOPEN1.Item("INV_BALANCE")
                        .Item("INV_PMT_CURR") = rowARTOPEN1.Item("INV_PMT")
                        .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW")
                    End With
                Next
            End If
        End If

        ASCMAIN1.Record_Event("SPTPYMT1", PYMT_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "P", "Payment Entered", "")

        Update_Record_TDA("SPTPYMT3")
        Update_Record_TDA("SPTPYMT2")
        Update_Record_TDA("SPTPYMT1")

        Update_Record_TDA("SPTCOOP1")

        If PYMT_TYPE = "R" Then
            Update_Record_TDA("ARTPYMT1")
            Update_Record_TDA("ARTPYMT2")
            Update_Record_TDA("ARTPYMT3")
            Update_Record_TDA("ARTOPEN1")
        Else
            Update_Record_TDA("APTINVH1")
            Update_Record_TDA("APTINVH2")
        End If

        If INVOICE_FROM_EMAIL <> "" Then
            Update_APTSUBM1(PYMT_CTL_NO)
            Update_Record_TDA("APTSUBM1")
        End If


        CommitTrans("Update Complete")

    End Sub

    'Sub Write_Audit(row As DataRow)

    '    Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
    '    With rowASTAUDT1
    '        .Item("TABLE_NAME") = "METPYMT2"
    '        .Item("KEY_VALUE") = row.Item("AUTH_NO")
    '        .Item("COLUMN_NAME") = "CLOSED"
    '        .Item("USER_ID") = ASCMAIN1.USER_ID
    '        .Item("INIT_DATE") = DATETIME_STAMP
    '        .Item("OLD_VALUE") = ""
    '        .Item("NEW_VALUE") = row.Item("AUTH_NO")
    '        .Item("FM_MODE") = DBNull.Value
    '        .Item("NOTES") = DBNull.Value
    '        .Item("KEY_VALUE2") = DBNull.Value
    '        .Item("KEY_LNO") = DBNull.Value
    '        .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
    '        .Item("SELECTION_NO") = Me.SELECTION_NO
    '        .Item("XNO") = Me.XNO
    '    End With

    '    dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
    'End Sub

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


    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("AUTH_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SPTPYMT1"
            E.COLUMN_NAME = "PYMT_NO"
            E.CODE_VALUE = Absx1.txtFor("PYMT_NO").Text
            If PYMT_TYPE = "R" Then
                E.DESC_VALUE = HFs("CUST_CODE")
            Else
                E.DESC_VALUE = HFs("VEND_CODE")
            End If
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "SPTPYMT1"
        E.TABLE_KEY_CAPTION = "Promo Event Spend Authorization"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PYMT_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("CUST_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Public Overrides Function Events_Context() As Events_Entity

        Dim E As New Events_Entity

        E.TABLE_NAME = "SPTPYMT1"
        E.TABLE_KEY_CAPTION = "Promo Event Payment"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PYMT_NO").Text
            If PYMT_TYPE = "R" Then
                E.TABLE_KEY_DESC = Absx1.txtFor("CUST_CODE").Text
            Else
                E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text
            End If
            E.TABLE_KEY_locked = True ' ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                ' sql_where = "VEND_TYPE = 'S'"

            Case "VEND_BUYER_CODE"
                sql_where = "VEND_BUYER_CODE in (Select USER_ID from ASTUSER1 where USER_STATUS = 'A')"

        End Select
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTCOOPX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTPYMT3, "S", "Show Filter")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                'For Each grow As UltraWinGrid.UltraGridRow In grdSPTCODE1.Rows
                '    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                '    grow.Update()
                'Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            'Case "Item Status Inquiry"
            '    Dim VEHICLE_CODE As String = grd.ActiveRow.Cells("VEHICLE_CODE").Text
            '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEHICLE_CODE)
            '    If rowSPTAVEH1 IsNot Nothing Then
            '        Context_Launch("View", VEHICLE_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        'Click_Command("New", e)
                    End If
                End If

            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        'Click_Command("New", e)
                    End If
                End If

            Case "PYMT_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode Then
                    'Click_Command("New")
                End If
            Case "VEND_CODE"
                If Not InquiryMode Then
                    'Click_Command("New")
                End If
            Case "PYMT_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
            Case "VEND_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PYMT_REF_AMT"
        End Select
    End Sub

#End Region

#Region "grdSPTPYMT2"

    Private Sub grdSPTPYMT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTPYMT2.AfterCellUpdate
        Select Case e.Cell.Column.Key

            Case "CLOSED"

                grdSPTPYMT2.ActiveRow.Update()

        End Select
    End Sub

    Private Sub grdSPTPYMT2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTPYMT2.AfterExitEditMode

        'Select Case grdSPTPYMT2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTPYMT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTPYMT2.AfterRowActivate
        With grdSPTPYMT2.DisplayLayout.Bands(0)
            If grdSPTPYMT2.ActiveRow.IsAddRow Then
                .Columns("VEHICLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTPYMT2.ActiveCell = grdSPTPYMT2.ActiveRow.Cells("VEHICLE_CODE")
                grdSPTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("VEHICLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTPYMT2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTPYMT2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdSPTPYMT2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSPTPYMT2.AfterRowUpdate
        DisplayTotals()
    End Sub

    Private Sub grdSPTPYMT2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTPYMT2.BeforeExitEditMode
        If grdSPTPYMT2.ActiveCell Is Nothing Then Exit Sub
        With grdSPTPYMT2.ActiveCell
            Select Case .Column.Key

            End Select
        End With
    End Sub

    Private Sub grdSPTPYMT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTPYMT2.BeforeRowUpdate
        With grdSPTPYMT2
            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_NO").Value = Absx1.CtlFor("PYMT_NO").Text
                    .ActiveRow.Cells("PYMT_LNO").Value = Val(dst.Tables("SPTPYMT2").Compute("Max(PYMT_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTPYMT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTPYMT2.ClickCellButton

        If grdSPTPYMT2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "VEHICLE_CODE"
            Case "LOCATION_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End Select
        grdClickCellButton(grdSPTPYMT2, sql_where, False)

    End Sub

    Private Sub grdSPTPYMT2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTPYMT2.Error
        grdSPTPYMT2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()

    End Sub

    Private Sub grdSPTCOOPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTCOOPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("PYMT_NO").Text = e.Row.Cells("PYMT_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Dim YP As String = cbeYP.Value
        Fill_Records("SPTCOOPX", YP)
        Sort_grdColumns(grdSPTCOOPX, "PYMT_NO".ToLower)
        grdSPTCOOPX.Text = "Entered in " & cbeYP.Text

        ASCMAIN1.sql = sqlAPTSUBM1 & " AND SUBMIT_STATUS = 'U'"
        Fill_Records("APTSUBM1", "", True, ASCMAIN1.sql)


    End Sub

    Sub Load_Chargebacks_Applied()

        Dim PYMT_BATCH_NO As String = rowSPTPYMT1.Item("PYMT_BATCH_NO") & ""
        If PYMT_BATCH_NO <> "" Then
            ASCMAIN1.sql = "Select ARTPYMT3.*, ARTOPEN1.INV_BALANCE INV_BALANCE_NOW" & vbCrLf _
                & " from ARTPYMT3,ARTPYMT2,ARTOPEN1" & vbCrLf _
                & " where ARTPYMT3.PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" & vbCrLf _
                & "   and ARTOPEN1.INV_TYPE = ARTPYMT3.INV_TYPE" & vbCrLf _
                & "   and ARTOPEN1.INV_NUM = ARTPYMT3.INV_NUM" & vbCrLf _
                & "   and ARTOPEN1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT3.PYMT_BATCH_LNO" & vbCrLf
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PYMT_BATCH_LNO")
                Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
                rowARTPYMT3.ItemArray = row.ItemArray
                rowARTPYMT3.Item("INV_BALANCE_NEW") = row.Item("INV_BALANCE_NOW")
                dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
            Next
            Sort_grdColumns(grdARTPYMT3, "PYMT_BATCH_ILNO")
        End If
    End Sub

    Sub Load_ChargeBacks()

        dst.Tables("ARTPYMT1").Rows.Clear()
        dst.Tables("ARTPYMT2").Rows.Clear()
        dst.Tables("ARTPYMT3").Rows.Clear()

        'ASCMAIN1.sql = "Select ARTOPEN1.* from ARTOPEN1" & vbCrLf _
        '    & " where CUST_CODE in " & vbCrLf _
        '    & " (Select Distinct CUST_CODE from ARTCUST1 where CUST_BILL_TO_CUST = '" & CUST_BILL_TO_CUST & "'" & vbCrLf _
        '    & " or CUST_CODE = '" & CUST_CODE & "' or CUST_CODE = '" & CUST_BILL_TO_CUST & "')" & vbCrLf _
        '    & " and INV_BALANCE <> 0"

        '     & " where INV_TYPE = 'B' and CUST_CODE in " & vbCrLf _

        ASCMAIN1.sql = "Select ARTOPEN1.* from ARTOPEN1" & vbCrLf _
           & " where CUST_CODE = '" & CUST_BILL_TO_CUST & "'" & vbCrLf _
           & "   and (INV_TYPE = 'B' or INV_TYPE = 'C')" & vbCrLf _
           & "   and INV_BALANCE <> 0"

        Dim PYMT_BATCH_ILNO As Integer = 0
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            PYMT_BATCH_ILNO += 1
            Write_ARTPYMT3_from_ARTOPEN1(row, PYMT_BATCH_ILNO)
        Next

        Sort_grdColumns(grdARTPYMT3, "PYMT_BATCH_ILNO")
    End Sub

    Function Write_ARTPYMT3_from_ARTOPEN1(row As DataRow, PYMT_BATCH_ILNO As Integer) As DataRow
        Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
        With rowARTPYMT3
            '"INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF", 
            For Each COLUMN_NAME As String In New String() _
                {"INV_TYPE", "INV_NUM", "REASON_CODE", "INV_DATE", "INV_DUE_DATE", "CUST_CODE_SO",
                 "CUST_STORE_NO", "INV_CUST_PO", "INV_BALANCE",
                  "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ORDR_TYPE_CODE", "CUST_CODE"}
                .Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
            Next
            .Item("PYMT_BATCH_NO") = "".PadLeft(10, "0")
            .Item("PYMT_BATCH_LNO") = 1
            .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
            .Item("INV_BALANCE_NEW") = .Item("INV_BALANCE")
            .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            '.Item("INV_PMT_CURR") = .Item("INV_PMT")
            '.Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
            '.Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
            .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW")
        End With

        dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)

        Return rowARTPYMT3
    End Function

    Private Sub optPYMT_TYPE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPYMT_TYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Toggle_PR()
    End Sub

    Sub Toggle_PR()
        Absx1.txtFor("VEND_CODE").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        lblVEND_CODE.Visible = (optPYMT_TYPE.Value = "P")
        lblCUST_CODE.Visible = (optPYMT_TYPE.Value = "R")
        lblCUST_BILL_TO_CUST.Visible = (optPYMT_TYPE.Value = "R")
        txtVEND_CODE.Visible = (optPYMT_TYPE.Value = "P")
        txtCUST_CODE.Visible = (optPYMT_TYPE.Value = "R")
        txtCUST_BILL_TO_CUST.Visible = (optPYMT_TYPE.Value = "R")
        txtVEND_NAME.Visible = (optPYMT_TYPE.Value = "P")
        txtCUST_NAME.Visible = (optPYMT_TYPE.Value = "R")
        tabDetails.Tabs("Customer Deductions Charged Back").Visible = (optPYMT_TYPE.Value = "R")
        SplitContainer2.Panel2Collapsed = (optPYMT_TYPE.Value = "R")
        grdAPTSUBM1.Selected.Rows.Clear()
    End Sub

    Private Sub grdSPTCOOP1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP1.ClickCellButton
        If grdSPTCOOP1.ActiveRow IsNot Nothing Then
            Dim AUTH_NO As String = grdSPTCOOP1.ActiveRow.Cells("AUTH_NO").Value
            If dst.Tables("SPTPYMT2").Select("AUTH_NO = '" & AUTH_NO & "'").Length = 0 Then
                Dim rowSPTCOOP1 As DataRow = dst.Tables("SPTCOOP1").Rows.Find(AUTH_NO)
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowSPTCOOP1.Item("CUST_CODE"))
                Dim rowSPTPYMT2 As DataRow = dst.Tables("SPTPYMT2").NewRow
                With rowSPTPYMT2
                    For Each DC As DataColumn In dst.Tables("SPTPYMT2").Columns
                        If dst.Tables("SPTCOOP1").Columns.Contains(DC.ColumnName) Then
                            .Item(DC.ColumnName) = rowSPTCOOP1.Item(DC.ColumnName)
                        End If
                    Next
                    .Item("PYMT_NO") = PYMT_NO
                    .Item("PYMT_LNO") = Val(dst.Tables("SPTPYMT2").Compute("MAX(PYMT_LNO)", "") & "") + 1
                    .Item("PYMT_REF_AMT") = rowSPTCOOP1.Item("OPEN_AMT")
                    .Item("AUTH_PNO") = Val(rowSPTCOOP1.Item("PYMTS") & "") + 1
                End With
                dst.Tables("SPTPYMT2").Rows.Add(rowSPTPYMT2)

                Fill_Records("SPTCOOP3", AUTH_NO)
                For Each rowSPTCOOP3 As DataRow In dst.Tables("SPTCOOP3").Select("")

                    Dim rowSPTPYMT3 As DataRow = dst.Tables("SPTPYMT3").NewRow
                    With rowSPTPYMT3
                        For Each COLUMN_NAME As String In New String() {"PYMT_NO", "PYMT_LNO", "AUTH_NO", "AUTH_PNO"}
                            rowSPTPYMT3.Item(COLUMN_NAME) = rowSPTPYMT2.Item(COLUMN_NAME)
                        Next
                        For Each COLUMN_NAME As String In New String() {"AUTH_LNO", "ITEM_CODE", "COLLECTION_CODE", "DIST_AMT"}
                            rowSPTPYMT3.Item(COLUMN_NAME) = rowSPTCOOP3.Item(COLUMN_NAME)
                        Next
                        Dim COLLECTION_CODE As String = rowSPTCOOP3.Item("COLLECTION_CODE") & ""
                        If COLLECTION_CODE <> "" Then
                            Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                            rowSPTPYMT3.Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
                        End If
                        rowSPTPYMT3.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
                    End With
                    dst.Tables("SPTPYMT3").Rows.Add(rowSPTPYMT3)
                Next

                grdSPTCOOP1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
            End If
        End If
    End Sub

    Private Sub grdSPTCOOP1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOP1.InitializeRow
        Dim AUTH_NO As String = e.Row.Cells("AUTH_NO").Value
        If dst.Tables("SPTPYMT2").Select("AUTH_NO = '" & AUTH_NO & "'").Length <> 0 Then
            e.Row.CellAppearance.BackColor = Drawing.Color.LightGreen
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If

        With e.Row.Cells("APPR_STATUS_CODE")
            Select Case .Value & ""
                Case "A"
                    .Appearance.ForeColor = Color.Green
                Case "P"
                    .Appearance.ForeColor = Color.Purple
                Case "G"
                    .Appearance.ForeColor = Color.Blue
                Case "X"
                    .Appearance.ForeColor = Color.Red

            End Select
        End With
    End Sub

    Private Sub grdARTPYMT3_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT3.AfterCellUpdate
        If e.Cell.Column.Key = "SELECTED" Then
            If e.Cell.Value = "1" Then
                e.Cell.Row.Cells("INV_PMT").Value = e.Cell.Row.Cells("INV_BALANCE").Value
            Else
                e.Cell.Row.Cells("INV_PMT").Value = 0
            End If
        End If
    End Sub

    Private Sub grdSPTPYMT2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTPYMT2.InitializeRow
        With e.Row.Cells("APPR_STATUS_CODE")
            Select Case .Value & ""
                Case "A"
                    .Appearance.ForeColor = Color.Green
                Case "P"
                    .Appearance.ForeColor = Color.Purple
                Case "G"
                    .Appearance.ForeColor = Color.Blue
                Case "X"
                    .Appearance.ForeColor = Color.Red
            End Select
        End With

        Dim PAID_AMT As Decimal = Val(e.Row.Cells("PAID_AMT").Value & "")
        Dim PYMT_REF_AMT As Decimal = Val(e.Row.Cells("PYMT_REF_AMT").Value & "")
        Dim TOTAL As Decimal = Val(e.Row.Cells("TOTAL").Value & "")

        Dim OVER As Decimal = PAID_AMT + PYMT_REF_AMT - TOTAL
        If OVER > SP_PARM_LIMIT_AMT Or (TOTAL = 0 OrElse 100 * OVER / TOTAL > SP_PARM_LIMIT_PCT) Then
            e.Row.Cells("PYMT_REF_AMT").Appearance.ForeColor = Color.Red
            e.Row.Cells("OPEN_AMT_CALC").Appearance.ForeColor = Color.Red
            e.Row.Cells("PYMT_REF_AMT").ToolTipText = "Payment exceeds overpayment limits"
        Else
            e.Row.Cells("PYMT_REF_AMT").Appearance.ForeColor = Color.Empty
            e.Row.Cells("OPEN_AMT_CALC").Appearance.ForeColor = Color.Empty
            e.Row.Cells("PYMT_REF_AMT").ToolTipText = ""
        End If
    End Sub

    Private Sub grdSPTCOOPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOPX.InitializeRow
        If e.Row.Cells("REVERSED_BY_PYMT_NO").Value & "" <> "" Or e.Row.Cells("REVERSED_PYMT_NO").Value & "" <> "" Then
            e.Row.Appearance.ForeColor = Color.Red
        Else
            e.Row.Appearance.ForeColor = Color.Empty
        End If
    End Sub

    Sub Print_Reconciling_Invoice()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing")

        Prepare_Invoice()

        Print_Report_Begin()
        ' CR_params.Add("NOTES", "1")
        Generate_Report("SPRPYMTA", "Promo Event Reconciling Invoice", "")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Prepare_Invoice()
        dst.Tables("SOTINVC1").Rows.Clear()
        dst.Tables("SOTINVC2").Rows.Clear()

        Fill_Records("ARTCUST1", CUST_BILL_TO_CUST)

        Dim INV_PMT As Decimal = Val(dst.Tables("ARTPYMT3").Compute("SUM(INV_PMT)", "PYMT_BATCH_ILNO <> 0") & "")
        Dim PYMT_REF_AMT As Decimal = Val(rowSPTPYMT1.Item("PYMT_REF_AMT") & "")

        Dim rowSOTINVC1 As DataRow = dst.Tables("SOTINVC1").NewRow
        With rowSOTINVC1
            .Item("PYMT_NO") = PYMT_NO
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            .Item("INV_TOTAL_AMOUNT") = INV_PMT - PYMT_REF_AMT
            .Item("PYMT_CTL_NO") = rowSPTPYMT1.Item("PYMT_CTL_NO")
            .Item("PYMT_REF_NO") = rowSPTPYMT1.Item("PYMT_REF_NO")
            .Item("PYMT_REF_AMT") = PYMT_REF_AMT
            .Item("INV_PMT") = INV_PMT
            .Item("CUST_STORE_NO") = ""
            .Item("SREP_CODE") = ""
            .Item("TERM_CODE") = ""
            .Item("PYMT_REF_DATE") = rowSPTPYMT1.Item("PYMT_REF_DATE")
            .Item("PYMT_REF_COMMENT") = "Pay up now, you deadbeat"
            .Item("AR_PARM_KEY") = "Z"
        End With
        dst.Tables("SOTINVC1").Rows.Add(rowSOTINVC1)


        Dim PYMT_LNO As Integer = 0

        For Each rowSPTPYMT2 As DataRow In dst.Tables("SPTPYMT2").Select("")
            Dim rowSOTINVC2 As DataRow = dst.Tables("SOTINVC2").NewRow
            With rowSOTINVC2
                .Item("PYMT_NO") = rowSPTPYMT1.Item("PYMT_NO")
                PYMT_LNO += 1
                .Item("PYMT_LNO") = PYMT_LNO
                .Item("PYMT_REF_AMT") = rowSPTPYMT2.Item("PYMT_REF_AMT")
                ' .Item("FEATURE_DESC") = rowSPTPYMT2.Item("FEATURE_DESC")
                .Item("BOOKING_NAME") = rowSPTPYMT2.Item("BOOKING_NAME")
                .Item("VEHICLE_CODE") = rowSPTPYMT2.Item("VEHICLE_CODE")
            End With
            dst.Tables("SOTINVC2").Rows.Add(rowSOTINVC2)
        Next

        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("PYMT_BATCH_ILNO > 0", "PYMT_BATCH_ILNO")

            'Dim PYMT_BATCH_NO As String = rowARTPYMT3.Item("PYMT_BATCH_NO")
            'Dim PYMT_BATCH_LNO As Integer = rowARTPYMT3.Item("PYMT_BATCH_LNO")
            'Dim rowARTPYMT2 As DataRow = LookUp("ARTPYMT2", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO})

            Dim INV_NUM As String = rowARTPYMT3.Item("INV_NUM")
            ASCMAIN1.sql = "Select ARTPYMT2.* from ARTPYMT5,ARTPYMT2" & vbCrLf _
                & " where ARTPYMT5.CHARGEBACK_NO = '" & INV_NUM & "'" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
                & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO"
            Dim rowARTPYMT2 As DataRow = ASCDATA1.GetDataRow

            Dim rowSOTINVC2 As DataRow = dst.Tables("SOTINVC2").NewRow
            With rowSOTINVC2
                .Item("PYMT_NO") = rowSPTPYMT1.Item("PYMT_NO")
                PYMT_LNO += 1
                .Item("PYMT_LNO") = PYMT_LNO
                .Item("INV_TYPE") = rowARTPYMT3.Item("INV_TYPE")
                .Item("INV_NUM") = rowARTPYMT3.Item("INV_NUM")
                .Item("INV_DATE") = rowARTPYMT3.Item("INV_DATE")
                .Item("CUST_STORE_NO") = rowARTPYMT3.Item("CUST_STORE_NO")
                .Item("INV_CUST_PO") = rowARTPYMT3.Item("INV_CUST_PO")
                .Item("INV_BALANCE") = rowARTPYMT3.Item("INV_BALANCE")
                .Item("INV_PMT") = rowARTPYMT3.Item("INV_PMT")
                If rowARTPYMT2 IsNot Nothing Then
                    .Item("CUST_PYMT_REF_NO") = rowARTPYMT2.Item("CUST_PYMT_REF_NO")
                    .Item("CUST_PYMT_REF_DATE") = rowARTPYMT2.Item("CUST_PYMT_REF_DATE")
                    .Item("CUST_PYMT_AMT") = rowARTPYMT2.Item("CUST_PYMT_AMT")
                End If
            End With
            dst.Tables("SOTINVC2").Rows.Add(rowSOTINVC2)
        Next
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing")
        Print_Document(PYMT_NO)

        'Print_Report_Begin()
        '' CR_params.Add("NOTES", "1")
        'Generate_Report("POROPRT1", "Purchase Order", "")
        'Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Print_Document(PYMT_NO As String, Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Dim REPORT As String = "SPRPYMT1"
        Dim RPT As String = REPORT

        If Not REPORTS.ContainsKey(REPORT) Then
            REPORTS.Add(REPORT, Load_rptClass(REPORT))
            REPORTS(REPORT).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORT).Fill_Records_RPT(New String() {" and PYMT_NO = '" & PYMT_NO & "'"})

        'To fill the report's dataset with data from this form's dataset:
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
        '        .dst.Tables(TABLE_NAME).Rows.Clear()
        '        Dim SQL As String = ""
        '        If TABLE_NAME = "SOTINVH1" Then
        '            SQL = "ORDR_NO = '" & ORDR_NO & "'"
        '        End If

        '        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(Sql)
        '            Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
        '            If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

        '                For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
        '                    Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
        '                    rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
        '                Next
        '            Else
        '                rowr.ItemArray = row.ItemArray
        '            End If
        '            .dst.Tables(TABLE_NAME).Rows.Add(rowr)
        '        Next
        '    Next
        '    .EnforceConstraints(True)
        'End With

        Dim REPORT_NO As String = ""

        With REPORTS(REPORT).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Promo Event Payment", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Promo Event Payment", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Private Sub grdARTPYMT3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMT3.InitializeRow
        If Val(e.Row.Cells("PYMT_BATCH_ILNO").Value & "") = 0 Then
            e.Row.Appearance.BackColor = Color.Yellow
        End If
    End Sub

    Sub Fill_SOTINVP0()
        Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
        With ROWs("ARTPARM1")
            rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
            rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            ' rowSOTINVP0.Item("REMIT3") = "Tel " & .Item("AR_PARM_REMIT_PHONE") & " Fax " & .Item("AR_PARM_REMIT_FAX")
            rowSOTINVP0.Item("REMIT3") = ""
            rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
        End With

        With ASCMAIN1.rowASTPARM1
            rowSOTINVP0.Item("ADDRESS0") = .Item("AS_PARM_INST_NAME") & ""
            rowSOTINVP0.Item("ADDRESS1") = .Item("AS_PARM_INST_ADDR1") & ""
            rowSOTINVP0.Item("ADDRESS2") = .Item("AS_PARM_INST_CITY") & ", " _
                    & .Item("AS_PARM_INST_STATE") & " " _
                    & .Item("AS_PARM_INST_ZIP_CODE") & " " _
                    & .Item("AS_PARM_INST_COUNTRY")

            Dim TEL As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & ""
            If TEL.Length = 10 Then
                TEL = "(" & Mid(TEL, 1, 3) & ")" & Mid(TEL, 4, 3) & "-" & Mid(TEL, 7, 4)
            End If
            Dim FAX As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""
            If FAX.Length = 10 Then
                FAX = "(" & Mid(FAX, 1, 3) & ")" & Mid(FAX, 4, 3) & "-" & Mid(FAX, 7, 4)
            End If
            rowSOTINVP0.Item("ADDRESS3") = "P " & TEL & " F " & FAX
            ' rowSOTINVP0.Item("ADDRESS3") = ""
        End With

        rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)

    End Sub

    Private Sub grdAPTSUBM1_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdAPTSUBM1.ClickCellButton
        If grdAPTSUBM1.ActiveCell.Column.Key = "EMAIL" Then
            Dim PEND_INVOICE_FILENAME = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_CTL_NO").Text & ".eml"
            If grdAPTSUBM1.ActiveRow.Cells("SUBMIT_NO_ORIG").Text <> "" Then
                PEND_INVOICE_FILENAME = grdAPTSUBM1.ActiveRow.Cells("SUBMIT_NO_ORIG").Text & ".eml"
            End If
            Show_Document(ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & PEND_INVOICE_FILENAME)
        End If
    End Sub
    Sub Update_APTSUBM1(VOUCHER_NO As String)

        Dim rowAPTSUBM1 As DataRow
        Dim SUBMIT_CTL_NO As String = INVOICE_FROM_EMAIL
        rowAPTSUBM1 = dst.Tables("APTSUBM1").Rows.Find(SUBMIT_CTL_NO)
        If rowAPTSUBM1 Is Nothing Then
        Else
            rowAPTSUBM1.Item("SUBMIT_CTL_NO") = SUBMIT_CTL_NO
            rowAPTSUBM1.Item("VOUCHER_NO") = VOUCHER_NO
            rowAPTSUBM1.Item("INV_NUM") = Absx1.txtFor("PYMT_REF_NO").Text
            rowAPTSUBM1.Item("INV_DATE") = Absx1.dteFor("PYMT_REF_DATE").Value
            rowAPTSUBM1.Item("INV_AMT") = Val(Absx1.numFor("PYMT_REF_AMT").Value & "") ' numPYMT_REF_AMT
            'rowAPTSUBM1.Item("LAST_DATE") = DATETIME_STAMP
            rowAPTSUBM1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowAPTSUBM1.Item("SUBMIT_STATUS") = "P"
            rowAPTSUBM1.Item("INV_SOURCE") = "P"
            rowAPTSUBM1.Item("INV_SOURCE_CTL_NO") = PYMT_NO
        End If

        Dim SUBMIT_SUBJECT As String = rowAPTSUBM1.Item("SUBMIT_SUBJECT") & ""

        Dim PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & SUBMIT_CTL_NO & ".eml"
        If rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & "" <> "" Then
            'PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & ".msg"
            PEND_INVOICE_FILENAME = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\" & rowAPTSUBM1.Item("SUBMIT_NO_ORIG") & ".eml"
        End If

        If PEND_INVOICE_FILENAME <> "" Then

            If Not dst.Tables.Contains("ASTATTA2") Then
                Create_TDA(dst.Tables.Add, "ASTATTA2", "*")
            End If
            Dim ATTACHMENT_NO As String = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")

            'If My.Computer.FileSystem.FileExists(PEND_INVOICE_FILENAME) Then
            'Dim ATTACH_PATH As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
            'My.Computer.FileSystem.CopyFile(PEND_INVOICE_FILENAME, ATTACH_PATH)



            Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
            With rowASTATTA2
                .Item("TABLE_NAME") = "SPTPYMT1"
                .Item("COLUMN_NAME") = "PYMT_NO"
                .Item("CODE_VALUE") = PYMT_NO
                .Item("ATTACHMENT_NO") = ATTACHMENT_NO
                .Item("ATTACHMENT_DESC") = SUBMIT_SUBJECT
                .Item("ATTACHMENT_FILENAME") = PEND_INVOICE_FILENAME
                .Item("ATTACHMENT_EXT") = "eml"
                .Item("COMPUTER_NAME") = My.Computer.Name
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ATTACHMENT_TYPE") = "EML"
                .Item("ATTACHMENT_ORIGINATOR") = ""
                .Item("ATTACHMENT_DATETIME") = DATETIME_STAMP
                .Item("ATTACHMENT_STATUS") = ""
                .Item("ATTACHMENT_NOTES") = ""
            End With
            dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
            Update_Record_TDA("ASTATTA2")
            dst.Tables("ASTATTA2").Rows.Clear()

            Try
                Dim ATTA_INVOICE_FILENAME As String = ASCMAIN1.Folders("Attach") & ATTACHMENT_NO
                My.Computer.FileSystem.CopyFile(PEND_INVOICE_FILENAME, ATTA_INVOICE_FILENAME, True)

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to copy email to Attachment")
            End Try


        End If

    End Sub

End Class