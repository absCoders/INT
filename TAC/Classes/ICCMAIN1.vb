Public Class ICCMAIN1

    Public Shared tblTasks As DataTable

    Public Shared Function Item_Cost_History(
    ByVal ITEM_CODE As String,
    ByVal MOS As Integer,
    Optional ByRef YP_start As String = "")

        If YP_start = "" Then YP_start = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * MOS)

        Dim CH() As Decimal
        ReDim CH(MOS)

        ASCMAIN1.sql = "Select GLTPARM2.OPS_YYYYPP, ICTCOSTA.ITEM_COST_TOTAL" _
        & " from ICTCOSTA,GLTPARM2" _
        & " where ICTCOSTA.ITEM_CODE (+) = '" & ITEM_CODE & "'" _
        & "   and ICTCOSTA.OPS_YYYYPP (+) = GLTPARM2.OPS_YYYYPP" _
        & "   and GLTPARM2.OPS_YYYYPP " _
        & " between '" & YP_start & "' and '" & ASCMAIN1.CYP & "'"
        Dim i As Integer = -1
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP DESC")
            i += 1
            CH(i) = Val(row.Item("ITEM_COST_TOTAL") & "")
        Next

        Return CH
    End Function

    Public Shared Sub Update_Return(ByVal frm As ASFBASE1)
        Dim rowSOTRTRN1 As DataRow = frm.dst.Tables("SOTRTRN1").Rows(0)
        frm.dst.Tables("ARTOPEN1").Rows.Clear()
        frm.dst.Tables("SOTINVH1").Rows.Clear()
        frm.dst.Tables("SOTINVH2").Rows.Clear()

        'Dim rowSOTREAS1 As DataRow = frm.dst.Tables("SOTREAS1").Rows.Find(New String() _
        '    {rowSOTRTRN1.Item("REASON_CODE")})
        'Dim ACCT_CODE_RTN As String = rowSOTREAS1.Item("ACCT_CODE") & ""

        Dim rowARTCUST1 As DataRow = frm.LookUp("ARTCUST1", rowSOTRTRN1.Item("CUST_CODE"))
        If rowARTCUST1.Item("CUST_BILL_TO_CUST") & "" = "" Then
            rowSOTRTRN1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE")
        Else
            rowSOTRTRN1.Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_BILL_TO_CUST")
        End If
        Dim rowARTCUST1_BT As DataRow = frm.LookUp("ARTCUST1", rowSOTRTRN1.Item("CUST_BILL_TO_CUST"))

        Dim rowARTPOST1 As DataRow = frm.LookUp("ARTPOST1", rowARTCUST1_BT.Item("POST_CODE"))
        Dim rowARTSTAX1 As DataRow = Nothing ' frm.LookUp("ARTSTAX1", rowARTCUST1_BT.Item("STAX_CODE") & "")
        Dim rowARTSTAX2 As DataRow = Nothing ' frm.LookUp("ARTSTAX2", frm.dst.Tables("SOTRTRN5").Rows(0).Item("CUST_ZIP_CODE") & "")
        'Dim rowSOTMISC1 As DataRow = frm.LookUp("SOTMISC1", frm.ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RSF") & "")

        Dim INV_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            'INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
            INV_NO = rowSOTRTRN1.Item("RTRN_NO")
        Else
            INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        End If
        rowSOTRTRN1.Item("INV_NO") = INV_NO

        Dim RTRN_GNO As Integer = 0
        Dim rowSOTRTRN3 As DataRow
        Dim DIST_AMT As Decimal = 0

        For Each rowSOTRTRN2 As DataRow In frm.dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows) ' ISNULL(RTRN_QTY_REF,0) <> RTRN_QTY
            Dim RTRN_QTY As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY") & "")
            Dim RTRN_QTY_REF As Int32 = 0 ' Val(rowSOTRTRN2.Item("RTRN_QTY_REF") & "") - THIS WOULD BE FOR A QTY REFUSED FOR CREDIT - TO BE RETURNED TO THE CUSTOMER

            Dim DIST_AMT_SALES As Decimal = (RTRN_QTY - RTRN_QTY_REF) * Val(rowSOTRTRN2.Item("RTRN_PRICE") & "")
            Dim DIST_AMT_COSTS As Decimal = (RTRN_QTY - RTRN_QTY_REF) * Val(rowSOTRTRN2.Item("ITEM_COST_STD") & "")

            Dim rowICTCLAS1 As DataRow = frm.dst.Tables("ICTCLAS1").Rows.Find(New String() {rowSOTRTRN2.Item("ITEM_CLASS_CODE")})

            For RTRN_GNO = 1 To 4
                If ((RTRN_GNO = 1 Or RTRN_GNO = 1) And DIST_AMT_SALES <> 0) _
                Or ((RTRN_GNO = 3 Or RTRN_GNO = 4) And DIST_AMT_COSTS <> 0) Then

                    ' DON'T DO AR HERE - DO IT ONCE FOR THE DOCUMENT

                    rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow

                    rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN2.Item("RTRN_NO")
                    rowSOTRTRN3.Item("RTRN_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                    rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO

                    rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                    If RTRN_GNO = 1 Then ' Sales Returns
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_SLS_RTN")
                        rowSOTRTRN3.Item("DIST_TYPE") = "SLSRTN"
                        DIST_AMT = DIST_AMT_SALES
                    ElseIf RTRN_GNO = 2 Then ' Accts Receivable
                        rowSOTRTRN3.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
                        'rowSOTRTRN3.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
                        'rowSOTRTRN3.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
                        'rowSOTRTRN3.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")
                        rowSOTRTRN3.Item("DIST_TYPE") = "AR"
                        DIST_AMT = -1 * DIST_AMT_SALES
                    ElseIf RTRN_GNO = 3 Then ' Inventory
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_ONH")
                        rowSOTRTRN3.Item("DIST_TYPE") = "INVTY"
                        DIST_AMT = DIST_AMT_COSTS
                    ElseIf RTRN_GNO = 4 Then  ' Cost of Goods Returned
                        rowSOTRTRN3.Item("ACCT_CODE") = rowICTCLAS1.Item("ACCT_CODE_CGS_RTN")
                        rowSOTRTRN3.Item("DIST_TYPE") = "CGR"
                        DIST_AMT = -1 * DIST_AMT_COSTS
                    End If
                    rowSOTRTRN3.Item("DIST_AMT") = Math.Round(DIST_AMT, 2)
                    frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)

                End If
            Next

            Dim rowSOTINVH2 As DataRow = frm.dst.Tables("SOTINVH2").NewRow
            rowSOTINVH2.Item("INV_TYPE") = "C"
            rowSOTINVH2.Item("INV_NO") = INV_NO
            rowSOTINVH2.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE") & String.Empty
            rowSOTINVH2.Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO") & String.Empty
            rowSOTINVH2.Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE") & String.Empty
            rowSOTINVH2.Item("INV_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
            rowSOTINVH2.Item("ITEM_CODE") = rowSOTRTRN2.Item("ITEM_CODE")
            rowSOTINVH2.Item("ORDR_UNIT_PRICE") = rowSOTRTRN2.Item("RTRN_PRICE")
            rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTRTRN2.Item("RTRN_PRICE")
            rowSOTINVH2.Item("ORDR_QTY_SHIP") = -1 * (RTRN_QTY - RTRN_QTY_REF)
            rowSOTINVH2.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
            rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = rowSOTRTRN2.Item("OPS_YYYYPP")
            rowSOTINVH2.Item("ITEM_UNIT_COST") = rowSOTRTRN2.Item("ITEM_COST_STD")
            frm.dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
        Next

        'Dim CUST_SHIP_TO_STATE As String = rowARTCUST1.Item("CUST_STATE") & ""
        'Dim CUST_SHIP_TO_ZIP_TAX As String = Mid((rowARTCUST1.Item("CUST_ZIP_CODE") & ""), 1, 5)
        'Dim STAX_EXEMPT As String = rowARTCUST1.Item("STAX_EXEMPT") & ""
        'If rowSOTRTRN1.Item("CUST_SHIP_TO_NO") & "" <> "" Then
        '    Dim rowARTCUST2 As DataRow = frm.LookUp("ARTCUST2", New String() _
        '    {rowSOTRTRN1.Item("CUST_CODE"), rowSOTRTRN1.Item("CUST_SHIP_TO_NO")})
        '    CUST_SHIP_TO_STATE = rowARTCUST2.Item("CUST_SHIP_TO_STATE") & ""
        '    CUST_SHIP_TO_ZIP_TAX = Mid((rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE") & ""), 1, 5)
        '    STAX_EXEMPT = rowARTCUST2.Item("STAX_EXEMPT") & ""
        'End If

        Dim rowSOTINVH1 As DataRow = frm.dst.Tables("SOTINVH1").NewRow
        rowSOTINVH1.Item("INV_TYPE") = "C"
        rowSOTINVH1.Item("INV_NO") = INV_NO
        rowSOTINVH1.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
        rowSOTINVH1.Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
        rowSOTINVH1.Item("ORDR_CUST_PO") = rowSOTRTRN1.Item("CUST_CLAIM_NO")
        rowSOTINVH1.Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
        rowSOTINVH1.Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
        rowSOTINVH1.Item("TERM_CODE") = frm.ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")
        rowSOTINVH1.Item("REASON_CODE") = rowSOTRTRN1.Item("REASON_CODE")
        rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTRTRN1.Item("CUST_BILL_TO_CUST")
        rowSOTINVH1.Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
        rowSOTINVH1.Item("INV_SALES") = -1 * Val(rowSOTRTRN1.Item("RTRN_SALES") & "")
        rowSOTINVH1.Item("INV_COGS") = -1 * Val(rowSOTRTRN1.Item("RTRN_COSTS") & "")
        rowSOTINVH1.Item("INV_FREIGHT") = -1 * Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & "")
        If Val(rowSOTRTRN1.Item("RTRN_HANDLING") & "") <> 0 Then
            rowSOTINVH1.Item("INV_MISC_CHG") = Val(rowSOTRTRN1.Item("RTRN_HANDLING") & "")
            'rowSOTINVH1.Item("MISC_CHG_CODE") = frm.ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RSF") & "" ' "RSF"
        End If
        rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = -1 * Val(rowSOTRTRN1.Item("RTRN_AMOUNT") & "")
        rowSOTINVH1.Item("INV_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
        rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = rowSOTRTRN1.Item("OPS_YYYYPP")
        rowSOTINVH1.Item("INIT_DATE") = frm.DATETIME_STAMP
        rowSOTINVH1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        'rowSOTINVH1.Item("INV_PRINTED") = "0"
        'rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = CUST_SHIP_TO_STATE
        'rowSOTINVH1.Item("CUST_SHIP_TO_ZIP_TAX") = CUST_SHIP_TO_ZIP_TAX
        'rowSOTINVH1.Item("STAX_CODE") = rowSOTRTRN1.Item("STAX_CODE")
        'rowSOTINVH1.Item("STAX_RATE") = 0
        'If rowARTSTAX2 IsNot Nothing And rowSOTRTRN1.Item("INV_NO_ORIG") & "" = "" Then
        'If rowARTSTAX1.Item("SALES_USE") & "" = "U" Then
        '    rowSOTINVH1.Item("STAX_RATE") = rowARTSTAX2.Item("COMBINED_USE_TAX")
        'Else
        '    rowSOTINVH1.Item("STAX_RATE") = rowARTSTAX2.Item("COMBINED_SALES_TAX")
        'End If
        'If rowSOTRTRN1.Item("INV_NO_ORIG") & "" = "" And (rowSOTRTRN1.Item("STAX_CODE") = "AR" Or rowSOTRTRN1.Item("STAX_CODE") = "IN") And ASCMAIN1.CYP <= "201104" Then
        '    rowSOTINVH1.Item("STAX_RATE") = 0
        'End If
        'rowSOTINVH1.Item("STAX_EXEMPT") = STAX_EXEMPT
        'If STAX_EXEMPT = "1" Then
        '    rowSOTINVH1.Item("STAX_RATE") = 0
        'End If
        'Else
        'If rowSOTRTRN1.Item("INV_NO_ORIG") & "" = "" Then
        '    rowSOTINVH1.Item("STAX_EXEMPT") = STAX_EXEMPT
        'Else 'Get rate from orig inv if it exists
        '    rowSOTINVH1.Item("STAX_RATE") = Val(ASCDATA1.GetDataValue("SELECT STAX_RATE FROM SOTINVH1 WHERE INV_NO=:PARM1", "V", New String() {rowSOTRTRN1.Item("INV_NO_ORIG")}) & "")
        '    rowSOTINVH1.Item("STAX_EXEMPT") = Val(ASCDATA1.GetDataValue("SELECT STAX_EXEMPT FROM SOTINVH1 WHERE INV_NO=:PARM1", "V", New String() {rowSOTRTRN1.Item("INV_NO_ORIG")}) & "")
        'End If
        'End If

        'rowSOTINVH1.Item("INV_STAX") = -1 * Val(rowSOTRTRN1.Item("TOTAL_STAX") & "")
        ' rowSOTINVH1.Item("REGISTER_IND") = "0" ' PROBABLY SHOULD RE-IMPLENT THIS

        rowSOTINVH1.Item("ORDR_TYPE_CODE") = "RTN"
        rowSOTINVH1.Item("CURR_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_CURR_CODE")
        rowSOTINVH1.Item("CURR_EXCH_RATE") = 1
        rowSOTINVH1.Item("INV_COMMENT") = rowSOTRTRN1.Item("RTRN_NOTE")

        'ORDR_NO()
        'ORDR_DATE_UPDATED()

        ' Ask Walter about these fields
        'rowSOTINVH1.Item("INV_SALES_CURR") = rowSOTINVH1.Item("INV_SALES")
        'rowSOTINVH1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT")
        'rowSOTINVH1.Item("INV_MISC_CHG_CURR") = rowSOTINVH1.Item("INV_MISC_CHG")
        ''rowSOTINVH1.Item("INV_STAX_CURR") = rowSOTINVH1.Item("INV_STAX")
        'rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")

        rowSOTINVH1.Item("SALES_DIVISION_CODE") = ""

        frm.dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

        Dim rowARTOPEN1 As DataRow = frm.dst.Tables("ARTOPEN1").NewRow
        ' "STAX_CODE","INV_STAX",
        For Each C As String In New String() _
        {"CUST_CODE", "INV_TYPE", "INV_DATE", "CUST_STORE_NO", "POST_CODE",
         "TERM_CODE", "SREP_CODE",
         "ORDR_NO", "INV_SALES", "INV_FREIGHT", "INV_TOTAL_AMOUNT",
         "REASON_CODE", "INIT_OPER", "INIT_DATE", "INV_MISC_CHG", "ORDR_TYPE_CODE", "SALES_DIVISION_CODE"}
            If (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") And C = "ORDR_NO" Then
                rowARTOPEN1.Item("INV_ORDR_NO") = rowSOTINVH1.Item(C)
            Else
                rowARTOPEN1.Item(C) = rowSOTINVH1.Item(C)
            End If
        Next

        ' Added 11/2/2015
        If rowSOTINVH1.Item("CUST_BILL_TO_CUST") & String.Empty <> String.Empty Then
            rowARTOPEN1.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_BILL_TO_CUST")
        End If

        rowARTOPEN1.Item("INV_TYPE") = "R"
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            rowARTOPEN1.Item("INV_NO") = rowSOTINVH1.Item("INV_NO")
        Else
            rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
        End If

        rowARTOPEN1.Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
        rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("CUST_CODE_SO") = rowSOTINVH1.Item("CUST_CODE")
        rowARTOPEN1.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowARTOPEN1.Item("CURR_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_CURR_CODE")
        rowARTOPEN1.Item("CURR_EXCH_RATE") = 1
        rowARTOPEN1.Item("INV_SALES_CURR") = rowARTOPEN1.Item("INV_SALES")
        rowARTOPEN1.Item("INV_DISC_CURR") = rowARTOPEN1.Item("INV_DISC")
        rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowARTOPEN1.Item("INV_FREIGHT")
        rowARTOPEN1.Item("INV_STAX_CURR") = rowARTOPEN1.Item("INV_STAX")
        rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowARTOPEN1.Item("INV_MISC_CHG")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowARTOPEN1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")
        rowARTOPEN1.Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
        rowARTOPEN1.Item("INV_NOTES") = rowSOTRTRN1.Item("RTRN_NOTE")
        frm.dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        Call frm.Update_Record_TDA("ARTOPEN1")
        Call frm.Update_Record_TDA("SOTINVH1")
        Call frm.Update_Record_TDA("SOTINVH2")

        'APPLY_TO_INV_TYPE(VARCHAR2(1))
        'APPLY_TO_INV_NO(VARCHAR2(10))

        ' Stop
        'TAC.SOCMAIN1.Update_Sales_History_Summary(rowSOTINVH1)

        frm.Update_Record_TDA("SOTRTRN1")
        frm.Update_Record_TDA("SOTRTRN2")

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select SOTRTRN1.OPS_YYYYPP, SOTRTRN1.WHSE_CODE" _
        & ", SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_QTY, 0 RTRN_QTY_REF " _
        & " from SOTRTRN2,SOTRTRN1 where SOTRTRN2.RTRN_NO = :PARM1" _
        & " and SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.RTRN_QTY,0) - NVL(R1.RTRN_QTY_REF,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND)" _
        & " VALUES (R1.ITEM_CODE, R1.WHSE_CODE, R1.RTRN_QTY-NVL(R1.RTRN_QTY_REF,0));" _
        & " END IF;" _
        & " UPDATE ICTSTAT1 " _
        & " SET WHSE_QTY_RTN = NVL(WHSE_QTY_RTN,0) + NVL(R1.RTRN_QTY,0) - NVL(R1.RTRN_QTY_REF,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE" _
        & " AND OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT1 (OPS_YYYYPP, ITEM_CODE, WHSE_CODE, WHSE_QTY_RTN)" _
        & " VALUES (R1.OPS_YYYYPP, R1.ITEM_CODE, R1.WHSE_CODE, R1.RTRN_QTY - NVL(R1.RTRN_QTY_REF,0));" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowSOTRTRN1.Item("RTRN_NO")})


        RTRN_GNO = 0

        DIST_AMT = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "") ' AR
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                rowSOTRTRN3.Item("ACCT_CODE") = rowARTPOST1.Item("POST_ACCT_RECV_ACCT")
            Else
                rowSOTRTRN3.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
            End If

            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "AR"
            rowSOTRTRN3.Item("DIST_AMT") = Math.Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        'DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_STAX") & "") ' Sales Tax
        'If DIST_AMT <> 0 Then
        '    rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
        '    rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
        '    rowSOTRTRN3.Item("RTRN_LNO") = 0
        '    RTRN_GNO += 1
        '    rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
        '    rowSOTRTRN3.Item("ACCT_CODE") = rowARTSTAX1.Item("ACCT_CODE")
        '    rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        '    rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        '    rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        '    rowSOTRTRN3.Item("DIST_TYPE") = "STAX"
        '    rowSOTRTRN3.Item("DIST_AMT") = MATH.ROUND(DIST_AMT, 2)
        '    frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        'End If

        DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_FREIGHT") & "") ' Freight
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
            rowSOTRTRN3.Item("ACCT_CODE") = frm.ROWs.Item("SOTPARM1").Item("SO_PARM_ACCT_FRT_INC") & ""
            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "FRT"
            rowSOTRTRN3.Item("DIST_AMT") = Math.Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        'SHOULD WE DO RSF AS AN ACCOUNT IN SOTPARM1 OR AS A MISC CHG CODE?
        ' CODE BELOW IS USING SOTPARM1, BUT ODG USES SOTMISC1 - ROWSOTMISC1 IS ALREADY SET UP ABOVE TO HANDLE THIS
        ' - LET'S WAIT UNTIL SOMEONE ASKS FOR IT
        DIST_AMT = -1 * Val(rowSOTINVH1.Item("INV_MISC_CHG") & "") ' Handling
        If DIST_AMT <> 0 Then
            rowSOTRTRN3 = frm.dst.Tables("SOTRTRN3").NewRow
            rowSOTRTRN3.Item("RTRN_NO") = rowSOTRTRN1.Item("RTRN_NO")
            rowSOTRTRN3.Item("RTRN_LNO") = 0
            RTRN_GNO += 1
            rowSOTRTRN3.Item("RTRN_GNO") = RTRN_GNO
            rowSOTRTRN3.Item("ACCT_CODE") = frm.ROWs.Item("SOTPARM1").Item("SO_PARM_ACCT_HND_FEE") & "" ' rowSOTMISC1.Item("ACCT_CODE")
            rowSOTRTRN3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowSOTRTRN3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowSOTRTRN3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowSOTRTRN3.Item("DIST_TYPE") = "HND"
            rowSOTRTRN3.Item("DIST_AMT") = Math.Round(DIST_AMT, 2)
            frm.dst.Tables("SOTRTRN3").Rows.Add(rowSOTRTRN3)
        End If

        frm.Update_Record_TDA("SOTRTRN3")

        ASCMAIN1.sql = "Update ICTITEM1 " _
            & " Set ITEM_ORDR_REL_CODE = null " _
            & " where ITEM_CODE in (Select ITEM_CODE from SOTRTRN2 where RTRN_NO = '" & rowSOTRTRN1.Item("RTRN_NO") & "' and RTRN_QTY_1 > 0)" _
            & "   and ITEM_ORDR_REL_CODE in ('S','R')"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Shared Sub Update_RTV(ByVal frm As ASFBASE1, ByVal rowICTIRTV1 As DataRow)

        Dim IC_PARM_ACCT_CODE_RTV_CLEARING As String = frm.ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_RTV_CLEARING")

        For Each rowICTIRTV2 As DataRow In frm.dst.Tables("ICTIRTV2").Select("", "", DataViewRowState.CurrentRows)
            Dim STANDARD_COST_TOTAL As Decimal = Val(rowICTIRTV2.Item("RTV_QTY") & "") * Val(rowICTIRTV2.Item("PRICE_CATGY_COST_TOTAL") & "")
            Dim PO_COST_TOTAL As Decimal = Val(rowICTIRTV2.Item("RTV_QTY") & "") * Val(rowICTIRTV2.Item("PO_COST") & "")
            Dim PURCHASE_PRICE_VARIANCE As Decimal = STANDARD_COST_TOTAL - PO_COST_TOTAL
            Dim rowICTCATG1 As DataRow = frm.dst.Tables("ICTCATG1").Rows.Find(New String() {rowICTIRTV2.Item("PROD_CATGY_CODE")})
            For RTV_GNO As Integer = 1 To 3
                Dim rowICTIRTV3 As DataRow = frm.dst.Tables("ICTIRTV3").NewRow
                rowICTIRTV3.Item("RTV_NO") = rowICTIRTV2.Item("RTV_NO")
                rowICTIRTV3.Item("RTV_LNO") = rowICTIRTV2.Item("RTV_LNO")
                rowICTIRTV3.Item("RTV_GNO") = RTV_GNO
                If RTV_GNO = 1 Then ' Inventory
                    rowICTIRTV3.Item("ACCT_CODE") = rowICTCATG1.Item("ACCT_CODE_INV")
                    rowICTIRTV3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowICTIRTV3.Item("DIST_TYPE") = "INVTY"
                    rowICTIRTV3.Item("DIST_AMT") = Math.Round(STANDARD_COST_TOTAL * -1, 2)
                ElseIf RTV_GNO = 2 Then ' RTV Clearing
                    rowICTIRTV3.Item("ACCT_CODE") = IC_PARM_ACCT_CODE_RTV_CLEARING
                    rowICTIRTV3.Item("SEG2_CODE") = frm.dst.Tables("ICTWHSE1").Rows(0).Item("SEG2_CODE")
                    rowICTIRTV3.Item("DIST_TYPE") = "RTV"
                    rowICTIRTV3.Item("DIST_AMT") = Math.Round(PO_COST_TOTAL, 2)
                Else 'Purchase Price Variance
                    rowICTIRTV3.Item("ACCT_CODE") = rowICTCATG1.Item("ACCT_CODE_PPV")
                    rowICTIRTV3.Item("SEG2_CODE") = frm.dst.Tables("ICTWHSE1").Rows(0).Item("SEG2_CODE")
                    rowICTIRTV3.Item("DIST_TYPE") = "PPV"
                    rowICTIRTV3.Item("DIST_AMT") = Math.Round(PURCHASE_PRICE_VARIANCE, 2)
                End If
                'rowICTIRTV3.Item("SEG2_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowICTIRTV3.Item("SEG3_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowICTIRTV3.Item("SEG4_CODE") = frm.ROWs.Item("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                frm.dst.Tables("ICTIRTV3").Rows.Add(rowICTIRTV3)
                rowICTIRTV2.Item("OPS_YYYYPP") = rowICTIRTV1.Item("OPS_YYYYPP")
                If PURCHASE_PRICE_VARIANCE = 0 And RTV_GNO = 2 Then
                    Exit For 'No entry for PPV if there is no variance
                End If
            Next
        Next

        Call frm.Update_Record_TDA("ICTIRTV1")
        Call frm.Update_Record_TDA("ICTIRTV2")
        Call frm.Update_Record_TDA("ICTIRTV3")



        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select ICTIRTV1.OPS_YYYYPP, ICTIRTV1.WHSE_CODE" _
        & ", ICTIRTV2.ITEM_CODE, ICTIRTV2.RTV_QTY " _
        & " from ICTIRTV2,ICTIRTV1 where ICTIRTV2.RTV_NO = :PARM1" _
        & " and ICTIRTV1.RTV_NO = ICTIRTV2.RTV_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD,0) + NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND, WHSE_QTY_HOLD)" _
        & " VALUES (R1.ITEM_CODE, R1.WHSE_CODE, 0, R1.RTV_QTY);" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowICTIRTV1.Item("RTV_NO")})

    End Sub

    Public Shared Function Finalize_RTV(ByVal RTV_NO As String, ByVal RANo As String, ByVal TrackingNo As String) As String

        Dim voucherNo As String = ASCDATA1.ExecuteSF("ICPIRTVF", New String() {"RTV_NO", "USER_ID", "RTV_RA_NO", "RTV_TRACKING_NO"}, New Object() {RTV_NO, ASCMAIN1.USER_ID, RANo, TrackingNo})

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
        & " Select ICTIRTV1.OPS_YYYYPP, ICTIRTV1.WHSE_CODE" _
        & ", ICTIRTV2.ITEM_CODE, ICTIRTV2.RTV_QTY " _
        & " from ICTIRTV2,ICTIRTV1 where ICTIRTV2.RTV_NO = :PARM1" _
        & " and ICTIRTV1.RTV_NO = ICTIRTV2.RTV_NO;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " UPDATE ICTSTAT2 " _
        & " SET WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) - NVL(R1.RTV_QTY,0)," _
        & " WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD,0) - NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND)" _
        & " VALUES (R1.ITEM_CODE, R1.WHSE_CODE, -1 * R1.RTV_QTY);" _
        & " END IF;" _
        & " UPDATE ICTSTAT1 " _
        & " SET WHSE_QTY_RTV = NVL(WHSE_QTY_RTV,0) + NVL(R1.RTV_QTY,0)" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE" _
        & " AND OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & " IF SQL%NOTFOUND THEN" _
        & " INSERT INTO ICTSTAT1 (OPS_YYYYPP, ITEM_CODE, WHSE_CODE, WHSE_QTY_RTV)" _
        & " VALUES (R1.OPS_YYYYPP, R1.ITEM_CODE, R1.WHSE_CODE, R1.RTV_QTY);" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {RTV_NO})

        Return voucherNo
    End Function

    Public Shared Sub Update_Adjustment(ByVal frm As ASFBASE0)

        frm.Update_Record_TDA("ICTIADJ1")
        frm.Update_Record_TDA("ICTIADJ2")

        For Each rowICTIADJ1 As DataRow In frm.dst.Tables("ICTIADJ1").Select("")
            Dim ADJ_NO_in As String = rowICTIADJ1.Item("ADJ_NO")
            ASCDATA1.ExecuteSP("ICPIADJI", "VN", New Object() {ADJ_NO_in, 1}, New String() {"ADJ_NO_in", "S"})
            ASCDATA1.ExecuteSP("ICPIADJG", "V", New Object() {ADJ_NO_in}, New String() {"ADJ_NO_in"})
        Next
    End Sub

    Public Shared Sub Update_WHTLOCBX(TRAN_TYPE As String, TRAN_NO As String)

        ASCDATA1.ExecuteSP("WHPLOCB2",
                           "VVV",
                           New String() {TRAN_TYPE, TRAN_NO, ASCMAIN1.SESSION_NO},
                           New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})
    End Sub

    Public Shared Sub Shuttle_ADJ_to_ICTTRAN1(ByVal frm As ASFBASE0)

        If frm.dst.Tables.Contains("ICTTRAN1") Then
            frm.dst.Tables("ICTTRAN1").Rows.Clear()
            frm.dst.Tables("ICTTRAN2").Rows.Clear()
        Else
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN1", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "ICTTRAN2", "*")
        End If

        For Each rowICTIADJ1 As DataRow In frm.dst.Tables("ICTIADJ1").Select("")
            Dim rowICTTRAN1 As DataRow = frm.dst.Tables("ICTTRAN1").NewRow
            With rowICTTRAN1
                .Item("OPS_YYYYPP") = rowICTIADJ1.Item("OPS_YYYYPP")
                .Item("TRAN_TYPE") = "A"
                .Item("TRAN_NO") = rowICTIADJ1.Item("ADJ_NO")
                .Item("TRAN_SOURCE_DOCUMENT") = ""
                .Item("TRAN_DATE") = rowICTIADJ1.Item("ADJ_DATE")
                .Item("TRAN_WHSE_CODE") = rowICTIADJ1.Item("WHSE_CODE")
                .Item("TRAN_ADJ_REASON_CODE") = rowICTIADJ1.Item("REASON_CODE")
                .Item("INIT_DATE") = rowICTIADJ1.Item("INIT_DATE")
                .Item("INIT_OPER") = rowICTIADJ1.Item("INIT_OPER")
                .Item("TRAN_STATUS_UPD") = "U"
                Dim rowICTWHSE1 As DataRow = frm.LookUp("ICTWHSE1", rowICTIADJ1.Item("WHSE_CODE"))
                .Item("TRAN_CCVRW_DESC") = rowICTWHSE1.Item("WHSE_DESC")
                .Item("TRAN_COMMENT") = rowICTIADJ1.Item("ADJ_NOTE")
            End With
            frm.dst.Tables("ICTTRAN1").Rows.Add(rowICTTRAN1)

            For Each rowICTIADJ2 As DataRow In frm.dst.Tables("ICTIADJ2").Select("ADJ_NO = '" & rowICTIADJ1.Item("ADJ_NO") & "'")
                Dim rowICTTRAN2 As DataRow = frm.dst.Tables("ICTTRAN2").NewRow
                With rowICTTRAN2
                    .Item("OPS_YYYYPP") = rowICTIADJ1.Item("OPS_YYYYPP")
                    .Item("TRAN_TYPE") = "A"
                    .Item("TRAN_NO") = rowICTIADJ2.Item("ADJ_NO")
                    .Item("TRAN_LNO") = rowICTIADJ2.Item("ADJ_LNO")
                    .Item("ITEM_CODE") = rowICTIADJ2.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = frm.LookUp("ICTITEM1", rowICTIADJ2.Item("ITEM_CODE"))
                    .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                    .Item("ITEM_UOM") = rowICTITEM1.Item("ITEM_UOM")
                    .Item("ITEM_COST") = rowICTIADJ2.Item("ITEM_COST_STD")
                    .Item("TRAN_QTY") = rowICTIADJ2.Item("ADJ_QTY")
                    .Item("TRAN_QTY2") = rowICTIADJ2.Item("ADJ_QTY")
                End With
                frm.dst.Tables("ICTTRAN2").Rows.Add(rowICTTRAN2)
            Next
        Next

        frm.Update_Record_TDA("ICTTRAN1")
        frm.Update_Record_TDA("ICTTRAN2")

        frm.dst.Tables("ICTTRAN1").Rows.Clear()
        frm.dst.Tables("ICTTRAN2").Rows.Clear()
    End Sub

    Public Shared Sub Update_Transfer(ByVal frm As ASFBASE0)

        frm.Update_Record_TDA("ICTIXFR1")
        frm.Update_Record_TDA("ICTIXFR2")

        For Each rowICTIXFR1 As DataRow In frm.dst.Tables("ICTIXFR1").Select("")
            Dim XFR_NO_in As String = rowICTIXFR1.Item("XFR_NO")
            ASCDATA1.ExecuteSP("ICPIXFRI", "VN", New Object() {XFR_NO_in, 1}, New String() {"XFR_NO_in", "S"})
            ASCDATA1.ExecuteSP("ICPIXFRG", "V", New Object() {XFR_NO_in}, New String() {"XFR_NO_in"})
        Next
    End Sub

    Public Shared Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String, ByVal TT As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Double
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))

        Dim YP As String = ""
        Dim TX As String = ""
        Dim TK As String = ""

        Select Case JOURNAL_TYPE

            Case "ICIA"
                TX = "ICTIADJ3"
                TK = "ADJ_NO"

            Case "ICIT"
                TX = "ICTIXFR3"
                TK = "XFR_NO"

            Case "ICIR"
                TX = "ICTIREC3"
                TK = "RECEIPT_NO"

            Case "ICIV"
                TX = "ICTIRTV3"
                TK = "RTV_NO"

        End Select

        ASCMAIN1.sql = "" _
        & " Select T1.OPS_YYYYPP, TX.ACCT_CODE" & vbCrLf _
        & ", TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE" & vbCrLf _
        & ", TX.DIST_TYPE, SUM (TX.DIST_AMT) DIST_AMT " & vbCrLf _
        & " from " & TX & " TX," & TT & " T1" & vbCrLf _
        & " where TX." & TK & " = T1." & TK & vbCrLf _
        & " group by T1.OPS_YYYYPP, TX.ACCT_CODE, TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE, TX.DIST_TYPE " & vbCrLf _
        & " order by T1.OPS_YYYYPP, TX.ACCT_CODE, TX.SEG2_CODE, TX.SEG3_CODE, TX.SEG4_CODE, TX.DIST_TYPE "

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")
            Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = row("OPS_YYYYPP")
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = row("ACCT_CODE")
            rowGLTINTF1("SEG2_CODE") = row("SEG2_CODE")
            rowGLTINTF1("SEG3_CODE") = row("SEG3_CODE")
            rowGLTINTF1("SEG4_CODE") = row("SEG4_CODE")
            'If ASCMAIN1.CLIENT = "INT" And JOURNAL_TYPE = "ICIA" Then
            '    rowGLTINTF1("SEG4_CODE") = "000"
            'End If
            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
            rowGLTINTF1("DETL_POSTING_AMT") = Math.Round(DETL_POSTING_AMT, 2)
            rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
            rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
            rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_NO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_NO") = DBNull.Value
            rowGLTINTF1("DETL_DESC") = DBNull.Value
            rowGLTINTF1("DETL_CVX_TYPE") = DBNull.Value
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Next

        Return JOURNAL_NO

    End Function

    Public Shared Sub Update_Receipt(ByVal frm As ASFBASE0, reversal_update As Boolean)

        ' Get a list of POs that are related to this receipt
        Dim PO_ORDER_NOs As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct(frm.dst.Tables("ICTIREC2"), New String() {"PO_ORDER_NO"}).Rows
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            PO_ORDER_NOs.Add(PO_ORDER_NO)
        Next

        ' Undo Commitments and Open PO status qtys
        For Each PO_ORDER_NO As String In PO_ORDER_NOs
            TAC.POCMAIN1.Production_Commit(-1, PO_ORDER_NO)
            TAC.POCMAIN1.ICTSTAT2_PO(-1, PO_ORDER_NO)
        Next

        ' Record the Component Consumption
        TAC.ICCMAIN1.Update_ICTIREC4(frm)

        If reversal_update Then
            ' Use ICTIREC4
        Else

            If frm.dst.Tables.Contains("ICT3PLTX") _
                And frm.dst.Tables("ICTIREC1").Rows(0).Item("WHSE_CODE") = "CLA" _
                And frm.dst.Tables("ICTIREC1").Rows(0).Item("RECEIPT_SOURCE") & "" = "A" Then
                frm.dst.Tables("ICTIREC4").Rows.Clear()
                Dim rowICTIREC2 As DataRow = frm.dst.Tables("ICTIREC2").Rows(0)
                Dim EXT_COST_MATLS As Decimal = 0

                Dim DTDATE As String = Split(frm.HFs("CLARINS_ASSY"), vbTab)(1)
                Dim DTTIME As String = Split(frm.HFs("CLARINS_ASSY"), vbTab)(2)
                Dim DTUSER As String = Split(frm.HFs("CLARINS_ASSY"), vbTab)(3)
                Dim sqly As String = "DTDATE = '" & DTDATE & "' AND DTTIME = '" & DTTIME & "' AND DTUSER = '" & DTUSER & "'"
                sqly &= " AND ITEM_CODE <> '" & rowICTIREC2.Item("ITEM_CODE") & "' AND DTTQTY < 0"
                For Each row As DataRow In frm.dst.Tables("ICT3PLTY").Select(sqly)

                    Dim DTADJC As String = row.GetParentRow("ICT3PLTX_ICT3PLTY").Item("DTADJC")
                    If DTADJC = "A21" Then ' Dis-Assembly - create adjustment
                        Record_Adjustment_for_DisAssembly(frm, row, frm.dst.Tables("ICTIREC1").Rows(0))

                    ElseIf DTADJC = "A03" Then ' Assembly - record consumption

                        Dim rowICTIREC4 As DataRow = frm.dst.Tables("ICTIREC4").Rows.Find _
                                                    (New Object() {rowICTIREC2.Item("RECEIPT_NO"), rowICTIREC2.Item("RECEIPT_LNO"), row.Item("ITEM_CODE")})
                        If rowICTIREC4 IsNot Nothing Then
                            rowICTIREC4.Item("QTY_CON") = Val(rowICTIREC4.Item("QTY_CON") & "") - 1 * Val(row.Item("DTTQTY") & "")
                            EXT_COST_MATLS += -1 * Val(row.Item("DTTQTY") & "") * Val(rowICTIREC4.Item("ITEM_COST_STD") & "")
                        Else
                            rowICTIREC4 = frm.dst.Tables("ICTIREC4").NewRow
                            rowICTIREC4.Item("RECEIPT_NO") = rowICTIREC2.Item("RECEIPT_NO")
                            rowICTIREC4.Item("RECEIPT_LNO") = rowICTIREC2.Item("RECEIPT_LNO")
                            rowICTIREC4.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                            Dim rowICTITEM1 As DataRow = frm.LookUp("ICTITEM1", rowICTIREC4.Item("ITEM_CODE"))
                            rowICTIREC4.Item("QTY_CON") = -1 * Val(row.Item("DTTQTY") & "")
                            rowICTIREC4.Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD")
                            EXT_COST_MATLS += Val(rowICTIREC4.Item("QTY_CON") & "") * Val(rowICTIREC4.Item("ITEM_COST_STD") & "")
                            rowICTIREC4.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
                            rowICTIREC4.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
                            rowICTIREC4.Item("LOCATION_CODE") = DBNull.Value
                            rowICTIREC4.Item("BAR_CODE") = DBNull.Value
                            rowICTIREC4.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                            frm.dst.Tables("ICTIREC4").Rows.Add(rowICTIREC4)
                        End If

                    End If
                Next
                rowICTIREC2.Item("EXT_COST_MATLS") = EXT_COST_MATLS
            End If
        End If

        ' Record the Receipt
        frm.Update_Record_TDA("ICTIREC1")
        frm.Update_Record_TDA("ICTIREC2")
        frm.Update_Record_TDA("ICTIREC4")
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        frm.Update_Record_TDA("ICTIRECL")

        frm.Update_Record_TDA("ICTIADJ1")
        frm.Update_Record_TDA("ICTIADJ2")

        ' the next few lines appear to want to update the adjustments one at a time,
        ' but the Update_Adjustments method does not accept the ADJ_NO for input
        ' - it updates all adjustments in ICTIADJ1/2, including Update_Record("ICTIADJ1/2")
        ' so if there are >1 adjustments, they will update ICTSTAT1/2 repeatedly with the same adjustments
        'For Each rowICTIADJ1 As DataRow In frm.dst.Tables("ICTIADJ1").Select("")
        '    With rowICTIADJ1
        '        Dim ADJ_NO_in As String = .Item("ADJ_NO")

        ICCMAIN1.Update_Adjustment(frm)
        '    End With
        'Next

        For Each rowICTIREC1 As DataRow In frm.dst.Tables("ICTIREC1").Select("")

            With rowICTIREC1
                Dim RECEIPT_NO_in As String = .Item("RECEIPT_NO")

                ASCDATA1.ExecuteSP("ICPIRECX", "VN", New Object() {RECEIPT_NO_in, 1}, New String() {"RECEIPT_NO_in", "S"})

                'Create_Accrual_FRT(RECEIPT_NO_in)
                'Create_Accrual_TRF(RECEIPT_NO_in)

                If reversal_update Then
                    ASCMAIN1.Record_Event("POTORDR1", .Item("PO_ORDER_NO"), "",
                        .Item("INIT_DATE"),
                        .Item("INIT_OPER"),
                        "REC-REV", "Receipt Reversal", RECEIPT_NO_in)
                Else
                    ASCMAIN1.sql = "Update ICTITEM1 " _
                        & " Set ITEM_ORDR_REL_CODE = null " _
                        & " where ITEM_CODE in (Select ITEM_CODE from ICTIREC2 where RECEIPT_NO = '" & RECEIPT_NO_in & "')" _
                        & "   and ITEM_ORDR_REL_CODE in ('S','R')"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.Record_Event("POTORDR1", .Item("PO_ORDER_NO"), "",
                        .Item("INIT_DATE"),
                        .Item("INIT_OPER"),
                        "REC", "Receipt", RECEIPT_NO_in)
                End If
            End With
        Next

        ' Redo Commitments and Open PO status qtys

        For Each PO_ORDER_NO As String In PO_ORDER_NOs
            Dim rowPOTORDR1 As DataRow = frm.LookUp("POTORDR1", PO_ORDER_NO)
            ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            frm.Fill_Records("POTORDR2", "", True, ASCMAIN1.sql)
            If rowPOTORDR1.Item("VEND_WHSE_CODE") & "" <> "" Then
                TAC.POCMAIN1.Update_POTORDR9(frm, PO_ORDER_NO, rowPOTORDR1.Item("VEND_WHSE_CODE"))
                ' TAC.POCMAIN1.Production_Commit(1, PO_ORDER_NO)
            End If
            TAC.POCMAIN1.ICTSTAT2_PO(1, PO_ORDER_NO)
        Next
    End Sub

    Public Shared Sub Create_Accrual_FRT(RECEIPT_NO_in As String)

        Dim sqlCOST_ACC As String = "ROUND (NVL(ICTFRTC1.FRT_CLASS_PCT_CUR * NVL(ICTIREC2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2)"
        ASCMAIN1.sql = "Insert into APTACRC1" & vbCrLf _
            & "Select TAPCTLN1('APTACRC1.CTL_NO',1) CTL_NO, NVL(APTVEND1.VEND_CODE_FRT,ICTIREC1.VEND_CODE) VEND_CODE_ACC, 'FRT' ACCRUAL_CODE, NULL COST_ACT" & vbCrLf _
            & $", {sqlCOST_ACC} COST_ACC" & vbCrLf _
            & ", NULL CHARGEBACK_IND, NULL VOUCHER_NO, '0' CTL_STATUS, ICTIREC1.OPS_YYYYPP" & vbCrLf _
            & ", NULL VOUCHER_NO_ORIG, ICTIREC1.RECEIPT_DATE CTL_DATE, NULL CTL_NOTE" & vbCrLf _
            & ", NULL INV_PRINT_IND, NULL INV_PRINT_DATE, NULL INV_PRINT_USER" & vbCrLf _
            & $", {sqlCOST_ACC} COST_ORIG" & vbCrLf _
            & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & ", ICTIREC2.ITEM_CODE, 'F' CTL_TYPE, ICTIREC2.COST_CATGY_CODE, ICTIREC1.SOURCE_DOC_NO, '0' PPD_MATCHED, '0' PPD_IND, NULL VAR_TOLERANCE" & vbCrLf _
            & ", NULL NOTES, NULL COST_VAR_ITEM, NULL PPD_MATCHED_XNO, ICTIREC1.FRT_CLASS_CODE XXX_CLASS_CODE, ICTFRTC1.FRT_CLASS_PCT_CUR XXX_CLASS_PCT" & vbCrLf _
            & ", NULL TPV_ADJ, NULL BOL_NO_MATCHED, NULL OPS_YYYYPP_MATCHED, NULL CTL_NO_MATCHED, NULL VAR_OK, NULL BOL_NO, NULL BOL_REVERSAL_IND" & vbCrLf _
            & " from ICTIREC2,ICTIREC1,ICTCOSTA,POTORDR1,POTORDR2,ICTFRTC1,APTVEND1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & $"   and ICTIREC2.RECEIPT_NO = '{RECEIPT_NO_in}'" & vbCrLf _
            & $"   and {sqlCOST_ACC} <> 0" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & "   and ICTFRTC1.FRT_CLASS_CODE (+) = ICTIREC1.FRT_CLASS_CODE" & vbCrLf _
            & "   and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"
        Dim RF As Integer = ASCDATA1.ExecuteSQL()

    End Sub

    Public Shared Sub Create_Accrual_TRF(RECEIPT_NO_in As String)

        Dim sqlCOST_ACC As String = "ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(ICTIREC2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2)"
        ASCMAIN1.sql = "Insert into APTACRC1" & vbCrLf _
            & "Select TAPCTLN1('APTACRC1.CTL_NO',1) CTL_NO, NVL(APTVEND1.VEND_CODE_TRF,ICTIREC1.VEND_CODE) VEND_CODE_ACC, 'TRF' ACCRUAL_CODE, NULL COST_ACT" & vbCrLf _
            & $", {sqlCOST_ACC} COST_ACC" & vbCrLf _
            & ", NULL CHARGEBACK_IND, NULL VOUCHER_NO, '0' CTL_STATUS, ICTIREC1.OPS_YYYYPP" & vbCrLf _
            & ", NULL VOUCHER_NO_ORIG, ICTIREC1.RECEIPT_DATE CTL_DATE, NULL CTL_NOTE" & vbCrLf _
            & ", NULL INV_PRINT_IND, NULL INV_PRINT_DATE, NULL INV_PRINT_USER" & vbCrLf _
            & $", {sqlCOST_ACC} COST_ORIG" & vbCrLf _
            & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & ", ICTIREC2.ITEM_CODE, 'T' CTL_TYPE, ICTIREC2.COST_CATGY_CODE, NVL(ICTPINV1.BOL_NO,ICTPINV1_REV.BOL_NO) SOURCE_DOC_NO, '0' PPD_MATCHED, '0' PPD_IND, NULL VAR_TOLERANCE" & vbCrLf _
            & ", NULL NOTES, NULL COST_VAR_ITEM, NULL PPD_MATCHED_XNO, ICTIREC1.TRF_CLASS_CODE XXX_CLASS_CODE, ICTTRFC1.TRF_CLASS_PCT_CUR XXX_CLASS_PCT" & vbCrLf _
            & ", NULL TPV_ADJ, NULL BOL_NO_MATCHED, NULL OPS_YYYYPP_MATCHED, NULL CTL_NO_MATCHED, NULL VAR_OK, NULL BOL_NO, NULL BOL_REVERSAL_IND" & vbCrLf _
            & " from ICTIREC2,ICTIREC1,ICTCOSTA,POTORDR1,POTORDR2,ICTTRFC1,ICTPINV1,ICTPINV2,APTVEND1" & vbCrLf _
            & ",ICTPINV1 ICTPINV1_REV,ICTPINV2 ICTPINV2_REV,ICTIREC2 ICTIREC2_REV,ICTIREC1 ICTIREC1_REV" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & $"   and ICTIREC2.RECEIPT_NO = '{RECEIPT_NO_in}'" & vbCrLf _
            & $"   and {sqlCOST_ACC} <> 0" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
            & "   and ICTTRFC1.TRF_CLASS_CODE (+) = ICTIREC1.TRF_CLASS_CODE" & vbCrLf _
            & "   and ICTPINV1.PINV_NO (+) = ICTPINV2.PINV_NO" & vbCrLf _
            & "   and ICTPINV2.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTPINV2.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & "   and ICTPINV1_REV.PINV_NO (+) = ICTPINV2_REV.PINV_NO" & vbCrLf _
            & "   and ICTPINV2_REV.RECEIPT_NO (+) = ICTIREC2_REV.RECEIPT_NO" & vbCrLf _
            & "   and ICTPINV2_REV.RECEIPT_LNO (+) = ICTIREC2_REV.RECEIPT_LNO" & vbCrLf _
            & "   and ICTIREC1_REV.RECEIPT_NO (+) = ICTIREC1.REVERSES_RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2_REV.RECEIPT_NO (+) = ICTIREC1.REVERSES_RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2_REV.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & "   and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"
        Dim RT As Integer = ASCDATA1.ExecuteSQL()



        '        ASCMAIN1.sql = "Select  'x' CTL_NO, NVL(APTVEND1.VEND_CODE_TRF,ICTIREC1.VEND_CODE) VEND_CODE_ACC, 'TRF' ACCRUAL_CODE, NULL COST_ACT" & vbCrLf _
        '& ", ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(POTORDR2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2) COST_ACC" & vbCrLf _
        '& ", NULL CHARGEBACK_IND, NULL VOUCHER_NO, '0' CTL_STATUS, ICTIREC1.OPS_YYYYPP" & vbCrLf _
        '& ", NULL VOUCHER_NO_ORIG, ICTIREC1.RECEIPT_DATE CTL_DATE, NULL CTL_NOTE" & vbCrLf _
        '& ", NULL INV_PRINT_IND, NULL INV_PRINT_DATE, NULL INV_PRINT_USER" & vbCrLf _
        '& ", ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(POTORDR2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2) COST_ORIG" & vbCrLf _
        '& ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
        '& ", ICTIREC2.ITEM_CODE, 'T' CTL_TYPE, ICTIREC2.COST_CATGY_CODE, ICTPINV1.BOL_NO SOURCE_DOC_NO, '0' PPD_MATCHED, '0' PPD_IND, NULL VAR_TOLERANCE" & vbCrLf _
        '& ", NULL NOTES, NULL COST_VAR_ITEM, NULL PPD_MATCHED_XNO, ICTIREC1.TRF_CLASS_CODE XXX_CLASS_CODE, ICTTRFC1.TRF_CLASS_PCT_CUR XXX_CLASS_PCT, NULL TPV_ADJ, NULL BOL_NO_MATCHED" & vbCrLf _
        '& " from ICTIREC2,ICTIREC1,ICTCOSTA,POTORDR1,POTORDR2,ICTTRFC1,ICTPINV1,ICTPINV2,APTVEND1" & vbCrLf _
        '& " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
        '& "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
        '& "   and ICTIREC2.RECEIPT_NO = '029550'" & vbCrLf _
        '& "   and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO" & vbCrLf _
        '& "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
        '& "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
        '& "   and ICTTRFC1.TRF_CLASS_CODE (+) = ICTIREC1.TRF_CLASS_CODE" & vbCrLf _
        '& "   and ICTPINV1.PINV_NO (+) = ICTPINV2.PINV_NO" & vbCrLf _
        '& "   and ICTPINV2.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
        '& "   and ICTPINV2.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
        '& "   and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"
        '        Dim tbl As DataTable = ASCDATA1.GetDataTable
        '        Dim r As Integer = tbl.Rows.Count
        '        Stop



        ' TRIED THIS, BUT ABANDONED IT BECAUSE
        ' CREATE UNIQUE INDEX I_APTACRC1_9 ON APTACRC1 (RECEIPT_NO, RECEIPT_LNO)
        ' ORA-01452: cannot CREATE UNIQUE INDEX; duplicate keys found
        ' MORE THAN LIKELY BECAUSE OF MANUAL ACCRUALS OR EARLY MIS-STEPS SINCE CORRECTED
        ' EXAMPLE: select * from APTACRC1 WHERE RECEIPT_NO = '026187' AND RECEIPT_LNO = 1 

        '        "Select 'X' CTL_NO, NVL(APTVEND1.VEND_CODE_TRF,ICTIREC1.VEND_CODE) VEND_CODE_ACC, 'TRF' ACCRUAL_CODE, NULL COST_ACT" & vbCrLf _
        '& ", ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(ICTIREC2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2) COST_ACC" & vbCrLf _
        '& ", NULL CHARGEBACK_IND, NULL VOUCHER_NO, '0' CTL_STATUS, ICTIREC1.OPS_YYYYPP" & vbCrLf _
        '& ", NULL VOUCHER_NO_ORIG, ICTIREC1.RECEIPT_DATE CTL_DATE, NULL CTL_NOTE" & vbCrLf _
        '& ", NULL INV_PRINT_IND, NULL INV_PRINT_DATE, NULL INV_PRINT_USER" & vbCrLf _
        '& ", ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(ICTIREC2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2) COST_ORIG" & vbCrLf _
        '& ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
        '& ", ICTIREC2.ITEM_CODE, 'T' CTL_TYPE, ICTIREC2.COST_CATGY_CODE, NVL(ICTPINV1.BOL_NO,APTACRC1_REV.SOURCE_DOC_NO) SOURCE_DOC_NO, '0' PPD_MATCHED, '0' PPD_IND, NULL VAR_TOLERANCE" & vbCrLf _
        '& ", NULL NOTES, NULL COST_VAR_ITEM, NULL PPD_MATCHED_XNO, ICTIREC1.TRF_CLASS_CODE XXX_CLASS_CODE, ICTTRFC1.TRF_CLASS_PCT_CUR XXX_CLASS_PCT, NULL TPV_ADJ, NULL BOL_NO_MATCHED, NULL OPS_YYYYPP_MATCHED" & vbCrLf _
        '& " from ICTIREC2,ICTIREC1,ICTCOSTA,POTORDR1,POTORDR2,ICTTRFC1,ICTPINV1,ICTPINV2,APTVEND1" & vbCrLf _
        '& ",ICTIREC1 ICTIREC1_REV,APTACRC1 APTACRC1_REV" & vbCrLf _
        '& " where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
        '& "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
        '& "   and ICTIREC2.RECEIPT_NO = '029985'" & vbCrLf _
        '& "   and ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(ICTIREC2.PO_COST,0)/100,0) * ICTIREC2.QTY_REC,2) <> 0" & vbCrLf _
        '& "   and POTORDR1.PO_ORDER_NO = ICTIREC1.PO_ORDER_NO" & vbCrLf _
        '& "   and POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO" & vbCrLf _
        '& "   and POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO" & vbCrLf _
        '& "   and ICTTRFC1.TRF_CLASS_CODE (+) = ICTIREC1.TRF_CLASS_CODE" & vbCrLf _
        '& "   and ICTPINV1.PINV_NO (+) = ICTPINV2.PINV_NO" & vbCrLf _
        '& "   and ICTPINV2.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
        '& "   and ICTPINV2.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
        '& "   and ICTIREC1_REV.RECEIPT_NO (+) = ICTIREC1.REVERSES_RECEIPT_NO" & vbCrLf _
        '& "   and APTACRC1_REV.RECEIPT_NO (+) = ICTIREC1.REVERSES_RECEIPT_NO" & vbCrLf _
        '& "   and APTACRC1_REV.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
        '& "   and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"
    End Sub

    Public Shared Sub Record_Adjustment_for_DisAssembly(ByVal frm As ASFBASE0, rowICT3PLTY As DataRow, rowICTIREC1 As DataRow)

        Dim WHSE_CODE As String = rowICTIREC1.Item("WHSE_CODE")

        Dim rowICTIADJ1 As DataRow = frm.dst.Tables("ICTIADJ1").NewRow
        rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        rowICTIADJ1.Item("WHSE_CODE") = WHSE_CODE
        rowICTIADJ1.Item("ADJ_DATE") = CDate(DateTime.Now.ToShortDateString)
        rowICTIADJ1.Item("ADJ_SOURCE") = "E"
        rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("INIT_DATE") = rowICTIREC1.Item("INIT_DATE")
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = rowICTIREC1.Item("INIT_DATE")
        rowICTIADJ1.Item("REGISTER_IND") = "0"
        rowICTIADJ1.Item("JOURNAL_IND") = "0"

        rowICTIADJ1.Item("REASON_CODE") = "A21"

        rowICTIADJ1.Item("ADJ_REF") = "" ' rowICTIREC1.Item("TRX_NO") & String.Empty

        Dim ITEM_CODE As String = rowICT3PLTY.Item("ITEM_CODE") & String.Empty
        Dim ADJ_NOTE As String = "Dis-Assembly of " & ITEM_CODE

        rowICTIADJ1.Item("ADJ_NOTE") = ADJ_NOTE

        frm.dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

        Dim ADJ_LNO As Int16 = 0

        Dim TOTAL_COSTS As Decimal = 0

        Dim rowICTITEM1 As DataRow = frm.LookUp("ICTITEM1", ITEM_CODE)
        Dim rowICTIADJ2 As DataRow = frm.dst.Tables("ICTIADJ2").NewRow

        rowICTIADJ2.Item("ADJ_NO") = rowICTIADJ1.Item("ADJ_NO")
        ADJ_LNO += 1
        rowICTIADJ2.Item("ADJ_LNO") = ADJ_LNO
        rowICTIADJ2.Item("ITEM_CODE") = ITEM_CODE
        rowICTIADJ2.Item("ADJ_QTY") = rowICT3PLTY.Item("DTTQTY")
        rowICTIADJ2.Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD") & String.Empty
        rowICTIADJ2.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE") & String.Empty
        rowICTIADJ2.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE") & String.Empty
        rowICTIADJ2.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIADJ2.Item("LOCATION_CODE") = "" ' LOCATION_CODE
        'rowICTIADJ2.Item("BAR_CODE") = String.Empty
        rowICTIADJ2.Item("ADJ_REF") = rowICTIADJ1.Item("ADJ_REF")
        frm.dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

        TOTAL_COSTS += Val(rowICTIADJ2.Item("ITEM_COST_STD") & String.Empty) * rowICTIADJ2.Item("ADJ_QTY")


        rowICTIADJ1.Item("TOTAL_COSTS") = TOTAL_COSTS
    End Sub

    Public Shared Sub Update_ICTIREC4(ByVal frm As ASFBASE0)
        For Each rowICTIREC1 As DataRow In frm.dst.Tables("ICTIREC1").Select("")
            If rowICTIREC1.Item("REVERSES_RECEIPT_NO") & "" <> "" Then
            Else
                Dim RECEIPT_NO As String = rowICTIREC1.Item("RECEIPT_NO")
                Dim sqlw As String = "RECEIPT_NO = '" & RECEIPT_NO & "' and VEND_WHSE_CODE is Not Null"
                For Each rowICTIREC2 As DataRow In frm.dst.Tables("ICTIREC2").Select(sqlw)
                    Dim PO_ORDER_NO As String = rowICTIREC2.Item("PO_ORDER_NO")
                    Dim PO_ORDER_LNO As Int32 = Val(rowICTIREC2.Item("PO_ORDER_LNO") & "")
                    Dim BM_ISSUE_SEL As String = rowICTIREC2.Item("BM_ISSUE_SEL") & ""
                    Dim BM_ISSUE_NO As String = rowICTIREC2.Item("BM_ISSUE_NO") & ""
                    Dim VEND_WHSE_CODE As String = rowICTIREC2.Item("VEND_WHSE_CODE")
                    Dim rowICTWHSE1 As DataRow = frm.LookUp("ICTWHSE1", VEND_WHSE_CODE)
                    Dim ITEM_CODE As String = rowICTIREC2.Item("ITEM_CODE")
                    Dim QTY_REC As Int64 = Val(rowICTIREC2.Item("QTY_REC") & "")
                    If QTY_REC <> 0 Then
                        Dim std_or_cur As String = ""
                        'Dim BM_ISSUE_NO As String = ""
                        If BM_ISSUE_SEL = "1" Then
                            std_or_cur = "C"
                            BM_ISSUE_NO = ""
                        Else
                            std_or_cur = ""
                            BM_ISSUE_NO = rowICTIREC2.Item("BM_ISSUE_NO") & ""
                        End If
                        If std_or_cur <> "" Or BM_ISSUE_NO <> "" Then
                            Dim TBL As DataTable = TAC.POCMAIN1.Get_BM(frm, ITEM_CODE, std_or_cur, BM_ISSUE_NO,
                                    False, True, "C", QTY_REC, VEND_WHSE_CODE, "HOPCA")
                            Dim EXT_COST_MATLS As Decimal = 0
                            For Each rowBMTMAIN3 As DataRow In TBL.Select("")
                                Dim rowICTIREC4 As DataRow = frm.dst.Tables("ICTIREC4").NewRow
                                rowICTIREC4.Item("RECEIPT_NO") = rowICTIREC2.Item("RECEIPT_NO")
                                rowICTIREC4.Item("RECEIPT_LNO") = rowICTIREC2.Item("RECEIPT_LNO")
                                rowICTIREC4.Item("ITEM_CODE") = rowBMTMAIN3.Item("BM_COMP_ITEM")
                                Dim rowICTITEM1 As DataRow = frm.LookUp("ICTITEM1", rowICTIREC4.Item("ITEM_CODE"))
                                rowICTIREC4.Item("QTY_CON") = rowBMTMAIN3.Item("QTY_COM")
                                rowICTIREC4.Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD")
                                EXT_COST_MATLS += Val(rowICTIREC4.Item("QTY_CON") & "") * Val(rowICTIREC4.Item("ITEM_COST_STD") & "")
                                rowICTIREC4.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
                                rowICTIREC4.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
                                rowICTIREC4.Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_SHP")
                                rowICTIREC4.Item("BAR_CODE") = DBNull.Value
                                rowICTIREC4.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                                frm.dst.Tables("ICTIREC4").Rows.Add(rowICTIREC4)
                            Next

                            If (BM_ISSUE_NO <> "" And rowICTIREC2.Item("BM_ISSUE_NO") & "" <> "") And
                               BM_ISSUE_NO <> rowICTIREC2.Item("BM_ISSUE_NO") & "" Then
                                'ASCDATA1.ExecuteSQL("Update POTORDR2 Set BM_ISSUE_NO = '" & BM_ISSUE_NO & "' where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO))
                            End If
                            rowICTIREC2.Item("BM_ISSUE_NO") = BM_ISSUE_NO
                            rowICTIREC2.Item("EXT_COST_MATLS") = EXT_COST_MATLS
                            'frm.Update_Record_TDA("ICTIREC2")
                        End If
                    End If
                Next
            End If
        Next
    End Sub

#Region "VB6"

    Public Shared Function ReCalculate_Costs(frmASFBASE0 As ASFBASE0, CFR As String, RE_VALUATION_IND As String) As Dictionary(Of String, String)

        Dim TABLES As New Dictionary(Of String, String)

        For Each TABLE_NAME As String In New String() {"ICTFRTC1", "ICTPROD1", "ICTCOST1", "TATCURR1", "ICTBRAN1", "ICTCOLL1"}
            If Not frmASFBASE0.dst.Tables.Contains(TABLE_NAME) Then
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, TABLE_NAME, "**", 0, False)
                frmASFBASE0.Fill_Records(TABLE_NAME)
            End If
        Next

        ASCMAIN1.Progress("-", "Item Master")

        ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_UOM, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ITEM_CATGY_CODE, PROD_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" & vbCrLf _
            & ", ITEM_COST_STATUS, ITEM_COST_MAKE_BUY, ITEM_CLASS_CODE" & vbCrLf _
            & ", ITEM_COST_FRT_CLASS, ITEM_COST_CURR_CODE" & vbCrLf _
            & ", ITEM_TYPE_CODE, COST_CATGY_CODE" & vbCrLf _
            & ", ITEM_COST_WASTE_PCT, ITEM_PLAN_WASTE_PCT" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, SALES_DIVISION_CODE, ITEM_COST_STD, 0 QTY_BOM" & vbCrLf _
            & ", ITEM_YYYYPP_CUR_COST, ITEM_YYYYPP_PRV_COST" & vbCrLf _
            & " from ICTITEM1, ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
        Dim ICTITEMX As String = ASCMAIN1.Temp_Table
        TABLES.Add("ICTITEMX", ICTITEMX)
        ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add Primary Key (ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Update " & ICTITEMX & " ICTITEMX Set ITEM_PLAN_WASTE_PCT = (Select ITEM_WASTE_PCT from ICTTYPE1 where ITEM_TYPE_CODE = ICTITEMX.ITEM_TYPE_CODE)")
        ASCDATA1.ExecuteSQL("Update " & ICTITEMX & " ICTITEMX Set ITEM_COST_WASTE_PCT = (Select ITEM_WASTE_PCT from ICTTYPE1 where ITEM_TYPE_CODE = ICTITEMX.ITEM_TYPE_CODE)")



        ASCMAIN1.Progress("-", "Future Cost Data")

        'ASCMAIN1.sql = "Select * from ICTCOSTF where ITEM_EXP_IMP_IND = 'E'"
        ASCMAIN1.sql = "Select * from ICTCOSTF where rownum<1"
        Dim ICTCOSTF As String = ASCMAIN1.Temp_Table
        TABLES.Add(" ICTCOSTF", ICTCOSTF)
        ASCDATA1.ExecuteSQL(" Alter Table " & ICTCOSTF & " Add Primary Key (ITEM_CODE)")


        ASCMAIN1.Progress("-", "Pending VCost Data")

        'ASCMAIN1.sql = "Select h.*, i.ITEM_DESC from ICTVCSTH h JOIN ICTITEM1 i ON h.ITEM_CODE=i.ITEM_CODE where ROWNUM < 1"
        'Dim ICTVCSTH As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        'TABLES.Add("ICTVCSTH", ICTVCSTH)

        ASCMAIN1.sql = Get_sqlICTCSTW1()

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM1", $"'{ASCMAIN1.CYP}'")
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM2", $"'{ASCMAIN1.CYP}'")
        'ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM3", $"'IPSA'")

        Dim ICTVCSTH As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        TABLES.Add("ICTVCSTH", ICTVCSTH)

        ASCDATA1.ExecuteSQL()

        'Dim ICTVCSTH As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql, frmASFBASE0.MENU_ITEM_OBJECT, frmASFBASE0.XNO, "VVV", New String() {ASCMAIN1.CYP, ASCMAIN1.CYP, "IPSA"})
        'Dim ICTVCSTH As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ' TABLES.Add("ICTVCSTH", ICTVCSTH)
        ASCDATA1.ExecuteSQL("Alter Table " & ICTVCSTH & " Add Primary Key (ITEM_CODE)")

        ' Insert or Update Future Cost Records based on Pending Cost Data
        ' Need to discuss if we have a place we are calculating freight?

        ASCMAIN1.sql = $"
        Begin Declare Cursor C1 is 
                Select h.*, COALESCE(t3.TRF_CLASS_PCT_FUT, t3.TRF_CLASS_PCT_CUR, t2.TRF_CLASS_PCT_FUT, t2.TRF_CLASS_PCT_CUR, t.TRF_CLASS_PCT_FUT, t.TRF_CLASS_PCT_CUR) TRF_PCT
        from {ICTVCSTH} h 
        join ictitem1 i on h.ITEM_CODE=i.ITEM_CODE 
        left join icttrfc1 t on i.item_cost_TRF_CLASS = t.TRF_CLASS_CODE
        left join {ICTCOSTF} f on h.item_code=f.item_code
        left join icttrfc1 t2 on f.ITEM_COST_TRF_CLASS=t2.TRF_CLASS_CODE
        left join icttrfc1 t3 on h.PEND_TRF_CLASS=t3.TRF_CLASS_CODE;
            Begin For R1 in C1 Loop
                Update {ICTCOSTF} Set ITEM_COST_VCOST = R1.PEND_VCOST, ITEM_COST_VCURR = R1.PEND_VCOST, ITEM_COST_TOOLG=round(R1.PEND_VCOST * nvl(R1.TRF_PCT,0)*.01,6)
                where ITEM_CODE = R1.ITEM_CODE;
            If SQL%NOTFOUND Then
                Insert into {ICTCOSTF} Select * from ICTCOSTC where ITEM_CODE = R1.ITEM_CODE;
                Update {ICTCOSTF} Set ITEM_COST_VCOST = R1.PEND_VCOST, ITEM_COST_VCURR = R1.PEND_VCOST, ITEM_COST_TOOLG=round(R1.PEND_VCOST * nvl(R1.TRF_PCT,0)*.01,6), ITEM_EXP_IMP_IND = 'E'
                where ITEM_CODE = R1.ITEM_CODE;
            End If;
            End Loop; End; 
        End;
        "
        ASCDATA1.ExecuteSQL()


        If ASCMAIN1.CYP = "202602" Then
            ASCMAIN1.sql = $"
            Begin Declare Cursor C1 is SELECT ICTITEM1.ITEM_CODE
FROM ICTCOSTF, ICTCOSTC, ICTITEM1
WHERE ICTITEM1.ITEM_CODE = ICTCOSTC.ITEM_CODE
AND ICTCOSTF.ITEM_CODE (+) = ICTCOSTC.ITEM_CODE
AND ICTITEM1.VEND_CODE = 'IPSA'
AND NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTC.ITEM_COST_TRF_CLASS) IN ('B');
             Begin For R1 in C1 Loop
              Update {ICTCOSTF} Set ITEM_COST_TRF_CLASS = 'A', ITEM_COST_TOOLG=round(ITEM_COST_VCOST * .10125,6)
                where ITEM_CODE = R1.ITEM_CODE;
              If SQL%NOTFOUND Then
               Insert into {ICTCOSTF} Select * from ICTCOSTC where ITEM_CODE = R1.ITEM_CODE;
               Update {ICTCOSTF} Set ITEM_COST_TRF_CLASS = 'A', ITEM_EXP_IMP_IND = 'E', ITEM_COST_TOOLG=round(ITEM_COST_VCOST * .10125,6)
                where ITEM_CODE = R1.ITEM_CODE;
              End If;
             End Loop; End; 
            End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"
            Begin Declare Cursor C1 is SELECT * FROM WJZ_NFC;
             Begin For R1 in C1 Loop
              Update {ICTCOSTF} Set ITEM_COST_VCOST = R1.STD_COST_NC, ITEM_COST_VCURR = R1.STD_COST_NC, ITEM_COST_TOOLG=round(R1.STD_COST_NC * DECODE(ITEM_COST_TRF_CLASS,'A',.10125,0),6)
                where ITEM_CODE = R1.ITEM_CODE;
              If SQL%NOTFOUND Then
               Insert into {ICTCOSTF} Select * from ICTCOSTC where ITEM_CODE = R1.ITEM_CODE;
               Update {ICTCOSTF} Set ITEM_COST_VCOST = R1.STD_COST_NC, ITEM_COST_VCURR = R1.STD_COST_NC, ITEM_EXP_IMP_IND = 'E', ITEM_COST_TOOLG=round(R1.STD_COST_NC * DECODE(ITEM_COST_TRF_CLASS,'A',.10125,0),6)
                where ITEM_CODE = R1.ITEM_CODE;
              End If;
             End Loop; End; 
            End;"
            ASCDATA1.ExecuteSQL()

        End If

        ASCMAIN1.Progress("-", "Current Cost Data")

        ASCMAIN1.sql = "Select * from ICTCOSTC"
        Dim ICTCOSTX As String = ASCMAIN1.Temp_Table
        TABLES.Add("ICTCOSTX", ICTCOSTX)
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add Primary Key (ITEM_CODE)")

        'ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTX & " where ITEM_CODE in (Select ITEM_CODE from ICTCOSTF where ITEM_EXP_IMP_IND = 'E')")
        ASCDATA1.ExecuteSQL($"Delete from {ICTCOSTX} where ITEM_CODE in (Select ITEM_CODE from {ICTCOSTF})")
        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " Set ITEM_EXP_IMP_IND = 'I'")
        'ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTX & " Select * from ICTCOSTF where ITEM_EXP_IMP_IND = 'E'")
        ASCDATA1.ExecuteSQL($"Insert into {ICTCOSTX} Select * from {ICTCOSTF}")

        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is Select * from ICTCOSTF where ITEM_EXP_IMP_IND = 'E';" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & "   Update " & ICTCOSTX & " ICTCOSTX " & vbCrLf _
        '    & "    Set ITEM_EXP_IMP_IND = R1.ITEM_EXP_IMP_IND, INIT_OPER = R1.INIT_OPER, INIT_DATE = R1.INIT_DATE" & vbCrLf _
        '    & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
        '    & "End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()

        Dim sqlCOLS As String = "ITEM_CODE, ITEM_COST_FRT_CLASS, ITEM_COST_MAKE_BUY, ITEM_CLASS_CODE, ITEM_CATGY_CODE, COST_CATGY_CODE"
        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTX & " (" & sqlCOLS & ") Select " & sqlCOLS & " from " & ICTITEMX & " where ITEM_CODE in (Select ITEM_CODE from " & ICTITEMX & " minus Select ITEM_CODE from " & ICTCOSTX & ")")


        ASCMAIN1.Progress("-", "Level Nesting")

        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " Set ITEM_LEVEL = '0'")
        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " Set ITEM_COST_MAKE_BUY = 'B' where ITEM_COST_MAKE_BUY is Null")

        ' THE FOLLOWING STMT OUGHT TO BE CONTROLLED BY A PARAMETER IN ICTPARM1
        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " Set ITEM_BM_ISSUE_SEL = 'C' where ITEM_COST_MAKE_BUY = 'M'")

        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " ICTCOSTX " & vbCrLf _
                            & " Set BM_ISSUE_NO = " & vbCrLf _
                            & "(Select Max (BM_ISSUE_NO) from BMTMAIN2 " & vbCrLf _
                            & " where BM_PROD_ITEM = ICTCOSTX.ITEM_CODE and BM_ISSUE_USE_FOR_STD = '1' and BM_ISSUE_NO <> '00')" & vbCrLf _
                            & " where ITEM_COST_MAKE_BUY = 'M' and ITEM_BM_ISSUE_SEL = 'C'")

        ' THE FOLLOWING STMT OUGHT TO BE CONTROLLED BY A PARAMETER IN ICTPARM1
        'ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " ICTCOSTX Set ITEM_COST_WASTE_PCT = (Select ITEM_PLAN_WASTE_PCT from ICTITEM1 where ITEM_CODE = ICTCOSTX.ITEM_CODE")
        ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " ICTCOSTX Set ITEM_COST_WASTE_PCT = (Select ITEM_PLAN_WASTE_PCT from " & ICTITEMX & " where ITEM_CODE = ICTCOSTX.ITEM_CODE)")

        ASCMAIN1.sql = "Select BM_PROD_ITEM, Max(BM_ISSUE_NO) BM_ISSUE_NO from BMTMAIN2 where BM_ISSUE_USE_FOR_STD = '1' group by BM_PROD_ITEM"
        ASCMAIN1.sql = "Select BMTMAIN3.* from BMTMAIN3 where (BM_PROD_ITEM,BM_ISSUE_NO) in (" & ASCMAIN1.sql & ") and NVL(BM_VEND_SUPP_MATL,'0') <> '1' and BM_WHEN_EXHAUSTED is Null"
        Dim BMTMAIN3 As String = ASCMAIN1.Temp_Table
        TABLES.Add("BMTMAIN3", BMTMAIN3)
        ASCDATA1.ExecuteSQL("Alter Table " & BMTMAIN3 & " Add Primary Key (BM_PROD_ITEM,BM_ISSUE_NO,BM_COMP_ITEM)")


        Dim II As Integer = ASCDATA1.ExecuteSQL("Update " & ICTCOSTX & " Set ITEM_LEVEL = 1 where ITEM_COST_MAKE_BUY = 'M'")
        If II > 0 Then
            Dim i As Integer = 0
            Do
                i += 1
                ASCMAIN1.sql = "Update " & ICTCOSTX & " set ITEM_LEVEL = '" & Format$(i + 1, "0") & "'" & vbCrLf _
                    & " where ITEM_LEVEL < '" & Format$(i + 1, "0") & "'" & vbCrLf _
                    & " and ITEM_CODE in " & vbCrLf _
                    & " (Select BM_PROD_ITEM from " & BMTMAIN3 & " where BM_COMP_ITEM in" & vbCrLf _
                    & " (Select ITEM_CODE from " & ICTCOSTX & " where ITEM_LEVEL = '" & Format$(i, "0") & "'))"
                II = ASCDATA1.ExecuteSQL
            Loop While II > 0
        End If


        ASCMAIN1.Progress("-", "Re-Calculate Costs")


        If Not frmASFBASE0.dst.Tables.Contains("ICTCOSTX") Then
            ASCMAIN1.sql = "Select * from " & ICTCOSTX
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTCOSTX"), ICTCOSTX, "**", 0, True)
        End If
        frmASFBASE0.Fill_Records("ICTCOSTX")
        If Not frmASFBASE0.dst.Tables.Contains("BMTMAIN3") Then
            ASCMAIN1.sql = "Select BMTMAIN3.* from BMTMAIN3" _
                & " where BMTMAIN3.BM_PROD_ITEM = :PARM1" & vbCrLf _
                & "   and BMTMAIN3.BM_ISSUE_NO = :PARM2" & vbCrLf _
                & "   and NVL(BMTMAIN3.BM_VEND_SUPP_MATL,'0') <> '1'" & vbCrLf _
                & "   and NVL(BMTMAIN3.BM_WHEN_EXHAUSTED,'?') = '?'"
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "BMTMAIN3", "**", 0, False, "VV")
            frmASFBASE0.Create_Relation("ICTCOSTX", "BMTMAIN3", "ITEM_CODE", "BM_COMP_ITEM")
            With frmASFBASE0.dst.Tables("BMTMAIN3").Columns
                .Add("ITEM_COST_WASTE_PCT", GetType(System.Decimal), "PARENT.ITEM_COST_WASTE_PCT")
                .Add("ITEM_COST_VCOST", GetType(System.Decimal), "PARENT.ITEM_COST_VCOST")
                .Add("ITEM_COST_LANDG", GetType(System.Decimal), "PARENT.ITEM_COST_LANDG")
                .Add("ITEM_COST_TOOLG", GetType(System.Decimal), "PARENT.ITEM_COST_TOOLG")
                .Add("ITEM_COST_OVRHD", GetType(System.Decimal), "PARENT.ITEM_COST_OVRHD")
                .Add("ITEM_COST_MATLS", GetType(System.Decimal), "PARENT.ITEM_COST_MATLS")
                .Add("ITEM_COST_LANDGI", GetType(System.Decimal), "PARENT.ITEM_COST_LANDGI")
                .Add("ITEM_COST_TOOLGI", GetType(System.Decimal), "PARENT.ITEM_COST_TOOLGI")
                .Add("ITEM_COST_OVRHDI", GetType(System.Decimal), "PARENT.ITEM_COST_OVRHDI")
                Dim M As String = " * ISNULL(BM_QTY_PER_ASSY,0) * (100 + ISNULL(ITEM_COST_WASTE_PCT,0)) / 100"
                .Add("EXT_MATLS", GetType(System.Decimal), "(ISNULL(ITEM_COST_VCOST,0)+ISNULL(ITEM_COST_MATLS,0))" & M)
                .Add("EXT_LANDGI", GetType(System.Decimal), "(ISNULL(ITEM_COST_LANDG,0)+ISNULL(ITEM_COST_LANDGI,0))" & M)
                .Add("EXT_TOOLGI", GetType(System.Decimal), "(ISNULL(ITEM_COST_TOOLG,0)+ISNULL(ITEM_COST_TOOLGI,0))" & M)
                .Add("EXT_OVRHDI", GetType(System.Decimal), "(ISNULL(ITEM_COST_OVRHD,0)+ISNULL(ITEM_COST_OVRHD,0))" & M)
            End With
        End If

        For Each rowICTCOSTX As DataRow In frmASFBASE0.dst.Tables("ICTCOSTX").Select("", "ITEM_LEVEL,ITEM_CODE")
            Dim ITEM_CODE As String = rowICTCOSTX.Item("ITEM_CODE")
            ASCMAIN1.Progress("-", ITEM_CODE)
            rowICTCOSTX.SetAdded()
            Calculate_Cost(frmASFBASE0, ITEM_CODE, ICTCOSTX, rowICTCOSTX)
            'If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH010C07USA" Then Stop
        Next


        ASCMAIN1.Progress("-", "Re-Loading Work Tables in Database")

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTCOSTX)
        frmASFBASE0.Update_Record_TDA("ICTCOSTX", "1=1")


        ASCMAIN1.Progress("-", "Preparing Re-Valuation Statistics")

        ASCMAIN1.sql = "Select ITEM_CODE, SUM (DECODE(WHSE_QTY_BEG,NULL,0,WHSE_QTY_BEG)) QTY_BEG from ICTSTAT1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' group by ITEM_CODE"
        Dim ICTSTATX As String = ASCMAIN1.Temp_Table
        TABLES.Add("ICTSTATX", ICTSTATX)
        ASCDATA1.ExecuteSQL("Alter Table " & ICTSTATX & " Add Primary Key (ITEM_CODE)")


        ' Clean out items w/no cost change, or cost status = 'P'

        If RE_VALUATION_IND = "R" Then
            ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTX & " where ITEM_CODE in (Select ITEM_CODE from " & ICTITEMX & " where ITEM_COST_STATUS = 'P')")

            ASCMAIN1.sql = "Select ICTCOSTX.ITEM_CODE from " & ICTCOSTX & " ICTCOSTX, ICTCOSTC " & vbCrLf _
            & " where ICTCOSTX.ITEM_CODE = ICTCOSTC.ITEM_CODE (+)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_VCOST,0)      = NVL(ICTCOSTC.ITEM_COST_VCOST,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_MATLS,0)      = NVL(ICTCOSTC.ITEM_COST_MATLS,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_LANDG,0)      = NVL(ICTCOSTC.ITEM_COST_LANDG,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_LANDGI,0)     = NVL(ICTCOSTC.ITEM_COST_LANDGI,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_TOOLG,0)      = NVL(ICTCOSTC.ITEM_COST_TOOLG,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_TOOLGI,0)     = NVL(ICTCOSTC.ITEM_COST_TOOLGI,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_OVRHD,0)      = NVL(ICTCOSTC.ITEM_COST_OVRHD,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_OVRHDI,0)     = NVL(ICTCOSTC.ITEM_COST_OVRHDI,0)" & vbCrLf _
            & "   and NVL(ICTCOSTX.ITEM_COST_TOTAL,0)      = NVL(ICTCOSTC.ITEM_COST_TOTAL,0)" & vbCrLf _
            & "   and ICTCOSTX.ITEM_COST_FRT_CLASS  = ICTCOSTC.ITEM_COST_FRT_CLASS" & vbCrLf _
            & "   and ICTCOSTX.ITEM_COST_MAKE_BUY   = ICTCOSTC.ITEM_COST_MAKE_BUY" & vbCrLf _
            & "   and ICTCOSTX.ITEM_COST_CURR_CODE  = ICTCOSTC.ITEM_COST_CURR_CODE" & vbCrLf _
            & "   and ICTCOSTX.ITEM_COST_WASTE_PCT  = ICTCOSTC.ITEM_COST_WASTE_PCT" & vbCrLf _
            & "   and ICTCOSTX.ITEM_COST_VCURR      = ICTCOSTC.ITEM_COST_VCURR" & vbCrLf _
            & "   and (ICTCOSTX.COLLECTION_CODE      = ICTCOSTC.COLLECTION_CODE" & vbCrLf _
            & "    or (ICTCOSTX.COLLECTION_CODE is Null and ICTCOSTC.COLLECTION_CODE is Null))" & vbCrLf _
            & "   and (ICTCOSTX.BM_ISSUE_NO          = ICTCOSTC.BM_ISSUE_NO" & vbCrLf _
            & "    or (ICTCOSTX.BM_ISSUE_NO is Null and ICTCOSTC.BM_ISSUE_NO is Null))" & vbCrLf _
            & "   and (ICTCOSTX.ITEM_COST_WASTE_TYPE = ICTCOSTC.ITEM_COST_WASTE_TYPE" & vbCrLf _
            & "    or (ICTCOSTX.ITEM_COST_WASTE_TYPE is Null and ICTCOSTC.ITEM_COST_WASTE_TYPE is Null))" & vbCrLf _
            & "   and (ICTCOSTX.ITEM_CLASS_CODE      = ICTCOSTC.ITEM_CLASS_CODE" & vbCrLf _
            & "    or (ICTCOSTX.ITEM_CLASS_CODE is Null and ICTCOSTC.ITEM_CLASS_CODE is Null))" & vbCrLf _
            & "   and (ICTCOSTX.COST_CATGY_CODE      = ICTCOSTC.COST_CATGY_CODE" & vbCrLf _
            & "    or (ICTCOSTX.COST_CATGY_CODE is Null and ICTCOSTC.COST_CATGY_CODE is Null))" & vbCrLf _
            & "   and (ICTCOSTX.ITEM_CATGY_CODE      = ICTCOSTC.ITEM_CATGY_CODE" & vbCrLf _
            & "    or (ICTCOSTX.ITEM_CATGY_CODE is Null and ICTCOSTC.ITEM_CATGY_CODE is Null))"
            ASCDATA1.ExecuteSQL("Delete from " & ICTCOSTX & " where ITEM_CODE in (" & ASCMAIN1.sql & ")")
            ASCDATA1.ExecuteSQL("Delete from " & ICTITEMX & " where ITEM_CODE in (Select ITEM_CODE from " & ICTITEMX & " minus Select ITEM_CODE from " & ICTCOSTX & ")")
            ASCDATA1.ExecuteSQL("Update " & ICTITEMX & " ICTITEMX Set QTY_BOM = (Select QTY_BEG from " & ICTSTATX & " where ITEM_CODE = ICTITEMX.ITEM_CODE)")
        End If

        frmASFBASE0.dst.Tables("BMTMAIN3").Rows.Clear()

        frmASFBASE0.Fill_Records("ICTCOSTX")

        If Not frmASFBASE0.dst.Tables.Contains("ICTITEMX") Then
            ASCMAIN1.sql = "Select * from " & ICTITEMX
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add("ICTITEMX"), ICTITEMX, "**", 0, False)
        End If
        frmASFBASE0.Fill_Records("ICTITEMX")

        If Not frmASFBASE0.dst.Tables.Contains("ICTCOSTC") Then
            ASCMAIN1.sql = "Select * from ICTCOSTC where ITEM_CODE in (Select ITEM_CODE from " & ICTCOSTX & ")"
            frmASFBASE0.Create_TDA(frmASFBASE0.dst.Tables.Add, "ICTCOSTC", "**", 0, False)
        End If
        frmASFBASE0.Fill_Records("ICTCOSTC")

        Return TABLES
    End Function

    Public Shared Sub Calculate_Cost(frmASFBASE0 As ASFBASE0, ITEM_CODE As String, ICTCOSTX As String, rowICTCOSTM As DataRow)
        Dim FRT_CLASS_CODE As String = rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") & ""
        Dim rowICTFRTC1 As DataRow = frmASFBASE0.dst.Tables("ICTFRTC1").Rows.Find(FRT_CLASS_CODE)
        Dim FRT_CLASS_PCT_CUR As Decimal = 0
        If rowICTFRTC1 IsNot Nothing Then
            FRT_CLASS_PCT_CUR = Val(rowICTFRTC1.Item("FRT_CLASS_PCT_CUR") & "")
        End If
        'If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH010C07USA" Then Stop
        rowICTCOSTM.Item("ITEM_COST_MATLS") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_LANDGI") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_TOOLGI") = DBNull.Value
        rowICTCOSTM.Item("ITEM_COST_OVRHDI") = DBNull.Value

        If rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "B" Then
        Else
            Dim BM_ISSUE_NO As String = rowICTCOSTM.Item("BM_ISSUE_NO") & ""
            frmASFBASE0.Fill_Records("BMTMAIN3", New String() {ITEM_CODE, BM_ISSUE_NO})
            Dim sqlw As String = "" ' "ISNULL(BM_WHEN_EXHAUSTED,'0') = '0'" - this was taken care if in calling programs
            rowICTCOSTM.Item("ITEM_COST_MATLS") = Val(frmASFBASE0.dst.Tables("BMTMAIN3").Compute("SUM(EXT_MATLS)", sqlw) & "")
            rowICTCOSTM.Item("ITEM_COST_LANDGI") = Val(frmASFBASE0.dst.Tables("BMTMAIN3").Compute("SUM(EXT_LANDGI)", sqlw) & "")
            rowICTCOSTM.Item("ITEM_COST_TOOLGI") = Val(frmASFBASE0.dst.Tables("BMTMAIN3").Compute("SUM(EXT_TOOLGI)", sqlw) & "")
            rowICTCOSTM.Item("ITEM_COST_OVRHDI") = Val(frmASFBASE0.dst.Tables("BMTMAIN3").Compute("SUM(EXT_OVRHDI)", sqlw) & "")
        End If

        Dim ITEM_COST_VCURR As Decimal = Val(rowICTCOSTM.Item("ITEM_COST_VCURR") & "")
        rowICTCOSTM.Item("ITEM_COST_VCOST") = ITEM_COST_VCURR
        rowICTCOSTM.Item("ITEM_COST_LANDG") = ITEM_COST_VCURR * FRT_CLASS_PCT_CUR / 100

        Dim ITEM_COST_TOTAL As Decimal = Val(rowICTCOSTM.Item("ITEM_COST_VCOST") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_LANDG") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_TOOLG") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_OVRHD") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_MATLS") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_LANDGI") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_TOOLGI") & "") _
                                       + Val(rowICTCOSTM.Item("ITEM_COST_OVRHDI") & "")

        rowICTCOSTM.Item("ITEM_COST_TOTAL") = ITEM_COST_TOTAL
        rowICTCOSTM.Item("ITEM_EXP_IMP_IND") = "E"

        Dim rowICTITEM1 As DataRow = frmASFBASE0.LookUp("ICTITEM1", ITEM_CODE)
        rowICTCOSTM.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
        rowICTCOSTM.Item("ITEM_CLASS_CODE") = rowICTITEM1.Item("ITEM_CLASS_CODE")
        rowICTCOSTM.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
        rowICTCOSTM.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
        rowICTCOSTM.Item("ITEM_COST_CURR_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        If frmASFBASE0.MENU_ITEM_OBJECT = "ICRRVAL1" Then
        Else
            rowICTCOSTM.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTCOSTM.Item("INIT_DATE") = frmASFBASE0.DATETIME_STAMP
        End If


        'BM_ISSUE_NO(VARCHAR2(2))
        'ITEM_COST_WASTE_TYPE(VARCHAR2(1))
        'ITEM_BM_ISSUE_SEL(VARCHAR2(1))
    End Sub

    Public Shared Sub Update_Movement_Costs()

        '  ICTSHIP2 (not yet)

        Dim X As String = "Set {0} = R1.ITEM_COST_STD where ITEM_CODE = R1.ITEM_CODE and {1} >= '" & ASCMAIN1.CYP & "';" & vbCrLf

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select ITEM_CODE, ITEM_COST_STD" & vbCrLf _
            & "  from ICTITEM1 where ITEM_YYYYPP_CUR_COST = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTINVH2 " & String.Format(X, "ITEM_UNIT_COST", "ORDR_YYYYPP_UPDATED") _
            & "   Update ICTIREC2 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "   Update ICTIREC4 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "   Update ICTIRTV2 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "   Update ICTIADJ2 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "   Update ICTIXFR2 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "   Update SOTRTRN2 " & String.Format(X, "ITEM_COST_STD", "OPS_YYYYPP") _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Update SOTINVH1 Set INV_COGS = (Select SUM (ORDR_QTY_SHIP * ITEM_UNIT_COST)" & vbCrLf _
            & "    from SOTINVH2 where INV_TYPE= SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)" & vbCrLf _
            & "   where ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & " End;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Declare Cursor C1 is Select * from ICTIRTV1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  Begin " & vbCrLf _
            & "   For R1 in C1 Loop" & vbCrLf _
            & "    ICPIRTVG(R1.RTV_NO);" & vbCrLf _
            & "   End Loop; " & vbCrLf _
            & "   Update ICTIRTV1 Set TOTAL_COSTS = (Select SUM (RTV_QTY * ITEM_COST_STD)" & vbCrLf _
            & "    from ICTIRTV2 where RTV_NO = ICTIRTV1.RTV_NO) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  End;" & vbCrLf _
            & " End;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Declare Cursor C1 is Select * from ICTIADJ1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  Begin " & vbCrLf _
            & "   For R1 in C1 Loop" & vbCrLf _
            & "    ICPIADJG(R1.ADJ_NO);" & vbCrLf _
            & "   End Loop; " & vbCrLf _
            & "   Update ICTIADJ1 Set TOTAL_COSTS = (Select SUM (ADJ_QTY * ITEM_COST_STD)" & vbCrLf _
            & "    from ICTIADJ2 where ADJ_NO = ICTIADJ1.ADJ_NO) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  End;" & vbCrLf _
            & " End;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Declare Cursor C1 is Select * from ICTIXFR1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  Begin" & vbCrLf _
            & "   For R1 in C1 Loop" & vbCrLf _
            & "    ICPIXFRG(R1.XFR_NO);" & vbCrLf _
            & "   End Loop;" & vbCrLf _
            & "   Update ICTIXFR1 Set TOTAL_COSTS = (Select SUM (XFR_QTY * ITEM_COST_STD)" & vbCrLf _
            & "    from ICTIXFR2 where XFR_NO = ICTIXFR1.XFR_NO) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  End;" & vbCrLf _
            & " End;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Declare Cursor C1 is Select * from ICTIREC1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  Begin" & vbCrLf _
            & "   Update ICTIREC2 Set EXT_COST_MATLS = (Select SUM (QTY_CON * ITEM_COST_STD)" & vbCrLf _
            & "    from ICTIREC4 where RECEIPT_NO = ICTIREC2.RECEIPT_NO AND RECEIPT_LNO = ICTIREC2.RECEIPT_LNO) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "   For R1 in C1 Loop" & vbCrLf _
            & "    ICPIRECG(R1.RECEIPT_NO);" & vbCrLf _
            & "    ICPIRECV(R1.RECEIPT_NO);" & vbCrLf _
            & "   End Loop; " & vbCrLf _
            & "   Update ICTIREC1 Set AMT_REC = (Select SUM (QTY_REC * ITEM_COST_STD)" & vbCrLf _
            & "    from ICTIREC2 where RECEIPT_NO = ICTIREC1.RECEIPT_NO) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  End;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Shared Function Check_Standard_Cost_Initialization(frmASFBASE0 As ASFBASE0, TABLE_NAME As String) As String

        Dim EMsg As String = ""

        For Each row As DataRow In frmASFBASE0.dst.Tables(TABLE_NAME).Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = frmASFBASE0.LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1.Item("ITEM_COST_STATUS") & "" = "P" Then
                EMsg &= vbCr & "Item " & ITEM_CODE & " is Pending Standard Cost Initialization"
            End If
        Next

        Return EMsg
    End Function

#End Region

    Public Shared Function Transmit_Document(
        MODULE_ID As String,
        MODULE_NAME As String,
        APP_CMD As String,
        APP_KEY As String,
        LP_CODE As String) As String
        ' Change for ADS 07/16/2025, force developer to supply the LP CODE

        Dim ErrorMessage As New List(Of String)
        Dim success As Boolean = False
        Return Transmit_Document(MODULE_ID, MODULE_NAME, APP_CMD, APP_KEY, ErrorMessage, success, LP_CODE)
    End Function

    Public Shared Function Transmit_Document(
        MODULE_ID As String,
        MODULE_NAME As String,
        APP_CMD As String,
        APP_KEY As String,
        ByRef ErrorMessages As List(Of String),
        ByRef Success As Boolean,
        LP_CODE As String) As String
        ' Change for ADS 07/16/2025, force developer to supply the LP CODE

        Dim sLocation As String = ""
        Dim buildType As String = "x86\Debug"

        If ASCMAIN1.Running_in_VS Then
            sLocation = ASCMAIN1.Folders("root") & MODULE_ID & "\bin\" & buildType & "\" & MODULE_ID & ".dll"
        Else
            sLocation = ASCMAIN1.Folders("bin") & MODULE_ID & ".dll"
        End If

        Dim sType As String = MODULE_ID & "." & MODULE_NAME

        Dim formAsm As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(sLocation)
        Dim ClassType As Type = formAsm.GetType(sType)

        Dim G As Object = Activator.CreateInstance(formAsm.GetType("WHC.ABSEnvironment"))
        G.DBS_COMPANY = ASCMAIN1.DBS_COMPANY
        G.DBS_SERVER = ASCMAIN1.DBS_SERVER
        G.DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD
        G.THREAD_NO = 0
        G.APP_ID = ""
        G.APP_DESC = ""
        G.USER_ID = ASCMAIN1.USER_ID
        G.APP_CMD = APP_CMD
        G.APP_KEY = APP_KEY
        G.LP_CODE = LP_CODE
        G.CLIENT = ASCMAIN1.CLIENT
        Dim C As Object = Activator.CreateInstance(ClassType, G)

        Try
            tblTasks = C.tblTasks.COPY
        Catch ex As Exception
        End Try

        ErrorMessages = C.ErrorMessages
        Success = C.SuccessfulExecution
        Return C.XMIT_NO
        ' C.Main_Process()

        ' AddHandler C.RespondToScan, AddressOf Display_Text
    End Function

    Public Shared Function Get_sqlICTCSTW1() As String
        '        Return "SELECT ICTCOSTP.OPS_YYYYPP as OPS_YYYYPP, ICTITEM1.ITEM_CODE, APTVEND1.VEND_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_STATUS, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_RETAIL_PRICE,
        'NVL(ICTCOSTF.ITEM_COST_VCOST,ICTCOSTC.ITEM_COST_VCOST) VCOST
        ', ICTCOSTP.ITEM_COST_VCOST CALC_VCOST -- will need to rename this column
        ', CASE WHEN ICTSTAT2_qualifies.ITEM_CODE IS NULL THEN 'N' ELSE 'Y' END AS QTY_OH_PO_IND
        ', CASE WHEN ICTSTAT1_qualifies.ITEM_CODE IS NULL THEN 'N' ELSE 'Y' END AS ACTIVITY_IND
        ', NULL as COST_LIST_PO_VCOST
        ', ICTITEM1.ITEM_DESC
        'FROM ICTCOSTF, ICTCOSTC, ICTITEM1, 
        '    (Select Distinct ITEM_CODE from ICTSTAT2
        '        where NVL(WHSE_QTY_ON_HAND,0) <> 0 OR NVL(WHSE_QTY_ONPO,0) <> 0) ICTSTAT2_qualifies,
        '    (Select Distinct ITEM_CODE from ICTSTAT1
        '        where OPS_YYYYPP = :PARM1
        '        and (NVL(WHSE_QTY_BEG,0) <> 0 or NVL(WHSE_QTY_SHP,0) <> 0 or NVL(WHSE_QTY_RTN,0) <> 0
        '        or NVL(WHSE_QTY_REC,0) <> 0 or NVL(WHSE_QTY_ADJ,0) <> 0 or NVL(WHSE_QTY_XFR,0) <> 0
        '        or NVL(WHSE_QTY_CON,0) <> 0 or NVL(WHSE_QTY_RTV,0) <> 0 or NVL(WHSE_QTY_PHY,0) <> 0
        '       )) ICTSTAT1_qualifies
        ', ICTCOSTP
        ', APTVEND1
        'WHERE ICTITEM1.ITEM_CODE = ICTCOSTC.ITEM_CODE
        'AND ICTCOSTF.ITEM_CODE (+) = ICTCOSTC.ITEM_CODE
        'AND ICTCOSTP.ITEM_CODE (+) = ICTITEM1.ITEM_CODE -- should this be a left join?
        'AND ICTSTAT2_qualifies.ITEM_CODE (+) = ICTITEM1.ITEM_CODE
        'AND ICTSTAT1_qualifies.ITEM_CODE (+) = ICTITEM1.ITEM_CODE
        'AND ICTITEM1.VEND_CODE = APTVEND1.VEND_CODE
        'AND ICTCOSTP.OPS_YYYYPP = :PARM2
        'AND ICTITEM1.VEND_CODE = :PARM3
        'ORDER BY ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_CODE"
        Return "SELECT ICTCOSTP.OPS_YYYYPP as OPS_YYYYPP, ICTITEM1.ITEM_CODE, ICTITEM1.VEND_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_STATUS, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_RETAIL_PRICE,
NVL(ICTCOSTF.ITEM_COST_VCOST,ICTCOSTC.ITEM_COST_VCOST) CUR_VCOST
, ICTCOSTP.ITEM_COST_VCOST PEND_VCOST
, NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTC.ITEM_COST_TRF_CLASS) CUR_TRF_CLASS
, ICTCOSTP.ITEM_COST_TRF_CLASS PEND_TRF_CLASS
, NVL(ICTCOSTF.ITEM_COST_FRT_CLASS,ICTCOSTC.ITEM_COST_FRT_CLASS) CUR_FRT_CLASS
, ICTCOSTP.ITEM_COST_FRT_CLASS PEND_FRT_CLASS
, CASE WHEN ICTSTAT2_qualifies.ITEM_CODE IS NULL THEN 'N' ELSE 'Y' END AS QTY_OH_PO_IND
, CASE WHEN ICTSTAT1_qualifies.ITEM_CODE IS NULL THEN 'N' ELSE 'Y' END AS ACTIVITY_IND
, ICTITEM1.ITEM_DESC
FROM ICTCOSTF, ICTCOSTC, ICTITEM1, 
    (Select Distinct ITEM_CODE from ICTSTAT2
        where NVL(WHSE_QTY_ON_HAND,0) <> 0 OR NVL(WHSE_QTY_ONPO,0) <> 0) ICTSTAT2_qualifies,
    (Select Distinct ITEM_CODE from ICTSTAT1
        where OPS_YYYYPP = :PARM1
        and (NVL(WHSE_QTY_BEG,0) <> 0 or NVL(WHSE_QTY_SHP,0) <> 0 or NVL(WHSE_QTY_RTN,0) <> 0
        or NVL(WHSE_QTY_REC,0) <> 0 or NVL(WHSE_QTY_ADJ,0) <> 0 or NVL(WHSE_QTY_XFR,0) <> 0
        or NVL(WHSE_QTY_CON,0) <> 0 or NVL(WHSE_QTY_RTV,0) <> 0 or NVL(WHSE_QTY_PHY,0) <> 0
       )) ICTSTAT1_qualifies
, ICTCOSTP
WHERE ICTITEM1.ITEM_CODE = ICTCOSTC.ITEM_CODE
AND ICTCOSTF.ITEM_CODE (+) = ICTCOSTC.ITEM_CODE
AND ICTCOSTP.ITEM_CODE (+) = ICTITEM1.ITEM_CODE
AND ICTSTAT2_qualifies.ITEM_CODE (+) = ICTITEM1.ITEM_CODE
AND ICTSTAT1_qualifies.ITEM_CODE (+) = ICTITEM1.ITEM_CODE
AND ICTCOSTP.OPS_YYYYPP = :PARM2
ORDER BY ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_CODE"
    End Function

End Class
