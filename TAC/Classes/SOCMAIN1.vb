Public Class SOCMAIN1

    Public Shared AGING_DATES() As Date

    Public Shared Function UPC( _
        ByRef frmASFBASE0 As ASFBASE0, _
        ByVal UPC_SEQUENCE_NO As String, _
        ByRef SO_PARM_UPC_VENDOR_ID As String, _
        Optional ByVal prefix_with_VENDOR_ID As Boolean = True) As String
        '
        'https://www.gs1.org/how-calculate-check-digit-manually

        ' Note: Check Digit Calculation applies to the 19-digits prior to the check digit
        '       These 11 digits are made up from the 6 digit Vendor ID prepended to the 5 digit UPC Serial Number
        '       19 digits = '0000' + 6 digit SO_PARM_UPC_VENDOR_ID + 9 digit Carton Serial Number

        Dim Check_Digit_Seed As String

        If prefix_with_VENDOR_ID Then
            If SO_PARM_UPC_VENDOR_ID = "" Then
                SO_PARM_UPC_VENDOR_ID = frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""
            End If

            If Len(UPC_SEQUENCE_NO) <> 5 Then
                If Len(UPC_SEQUENCE_NO) <> 9 Then
                    Stop
                End If
            End If

            Check_Digit_Seed = Mid(SO_PARM_UPC_VENDOR_ID, 1) & UPC_SEQUENCE_NO
        Else
            Check_Digit_Seed = UPC_SEQUENCE_NO
        End If

        Dim odd_digits As Integer
        Dim even_digits As Integer

        For i As Integer = 1 To Len(Check_Digit_Seed) Step 2
            odd_digits = odd_digits + Val(Mid(Check_Digit_Seed, 1, 1))
            Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            If Check_Digit_Seed <> "" Then
                even_digits = even_digits + Val(Mid(Check_Digit_Seed, 1, 1))
                Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            End If
        Next i

        Dim check_digit As Integer
        check_digit = (odd_digits * 3 + even_digits) Mod 10
        If check_digit <> 0 Then
            check_digit = 10 - check_digit
        End If

        If prefix_with_VENDOR_ID Then
            UPC = SO_PARM_UPC_VENDOR_ID & UPC_SEQUENCE_NO & Format(check_digit, "0")
        Else
            UPC = UPC_SEQUENCE_NO & Format(check_digit, "0")
        End If

    End Function

    Public Shared Function Credit_Check( _
    ByVal rowSOTORDR1 As DataRow, _
    ByVal rowARTCREDC As DataRow, _
    ByVal ORDR_AMT_OPEN As Decimal, _
    ByVal rowSOTPARM1 As DataRow, _
    Optional ByVal ARTCREDC As String = "", _
    Optional ByVal sqlARTCREDC As String = "") As String

        ' REVIEW THESE CHANGES WITH ED
        ' USE AS A DEMO OF WHEN A CLASS SHOULD BE INSTANTIATED
        ' GET CLS OUT OF THIS COLLECTION IN TATTERM1

        'TERM_C TERM_DESC                          
        '------ -----------------------------------
        'C1     COD CERTIFI CHK                    
        'C2     Credit On Acct                     
        'C3     Credit Card                        
        'C4     COD COMPANY CHK                    
        'P9     WEB CREDIT CARD                    

        Dim TERM_CODE_C As New List(Of String)
        TERM_CODE_C.Add("C1")
        TERM_CODE_C.Add("C2")
        TERM_CODE_C.Add("C3")
        TERM_CODE_C.Add("C4")
        TERM_CODE_C.Add("P9")

        Dim credit_card_auth As Boolean = False
        Dim C1_C2_C3_P9 As Boolean = False
        Dim past_due As Boolean = False

        Dim CCPA_NO As String = rowSOTORDR1.Item("CCPA_NO") & String.Empty

        If rowSOTORDR1.Item("ORDR_HOLD_CREDIT_REL_BY") & String.Empty <> "" Then
            Return String.Empty ' Order was Released by Credit Dept
        End If
        If rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty = "B2C" Then
            Return String.Empty ' B2C Orders are pre-paid by CC
        End If

        Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE") & String.Empty

        If Val(rowSOTORDR1.Item("ORDR_COD_ADDON_AMT") & "") <> 0 _
        Or rowSOTORDR1.Item("ORDR_HOLD_CREDIT_SPECIAL") & "" = "1" _
        Then
            ' order must visit the credit release screen
        Else

            If TERM_CODE_C.Contains(TERM_CODE) And TERM_CODE <> "C4" Then ' COMPANY COD CHECK MUST STILL BE WITHIN CREDIT LIMITS AND CREDIT HOLD CRITERIA
                'Return String.Empty ' Terms are a CC Type or a COD Type - Order is Pre-paid prior to Shipment
                C1_C2_C3_P9 = True
            End If

            If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then ' If Order is to be paid by CC and the CC was Authorized, then let it go
                'Return String.Empty
                credit_card_auth = True
            End If
        End If

        If ARTCREDC <> "" Then

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")

            ' SIMILAR CODE EXISTS IN SOFPICK1
            Dim SOTPICKC As String = "(" _
            & "Select SOTORDR1.CUST_CODE" _
            & ", Sum (NVL(SOTPICK2.PICK_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_PT" _
            & " from SOTPICK2,SOTPICK1,SOTORDR1,SOTORDR2" _
            & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO " _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
            & "   and (SOTPICK1.PICK_STATUS = 'N' or SOTPICK1.PICK_STATUS = 'P')" _
            & "   and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'" _
            & "   and SOTORDR1.ORDR_STATUS = 'P'" _
            & " group by SOTORDR1.CUST_CODE" _
            & ")"


            If AGING_DATES Is Nothing Then
                Initialize_AGING_DATES()
            End If

            ' SIMILAR CODE EXISTS IN SOFPICK1
            Dim ARTOPENA As String = "(" _
            & " Select CUST_CODE" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(1), "dd-MMM-yyyy") & "'                                 THEN INV_BALANCE ELSE 0 END) AGE_1" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(2), "dd-MMM-yyyy") & "' AND INV_DATE <= '" & Format(AGING_DATES(1), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_2" _
            & ", SUM (CASE WHEN INV_DATE > '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' AND INV_DATE <= '" & Format(AGING_DATES(2), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_3" _
            & ", SUM (CASE WHEN INV_DATE                                                                <= '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' THEN INV_BALANCE ELSE 0 END) AGE_4" _
            & ", SUM (CASE WHEN INV_DATE                                                                <= '" & Format(AGING_DATES(3), "dd-MMM-yyyy") & "' AND INV_TYPE IN ('I','D','B','O','C','R') THEN INV_BALANCE ELSE 0 END) PAST_DUE_DR_AMT" _
            & ", SUM (INV_BALANCE) TOTAL_DUE" _
            & " from ARTOPEN1 where INV_BALANCE <> 0" _
            & " and CUST_CODE = '" & CUST_CODE & "'" _
            & " group by CUST_CODE" _
            & ")"

            'Prepare_ARTCREDC(ARTCREDC, Replace(sqlARTCREDC, ":PARM1", "'" & CUST_CODE & "'"), SOTPICKC, "", ARTOPENA, "")
            'rowARTCREDC = ASCDATA1.GetDataRow("Select * from " & ARTCREDC)

            ASCMAIN1.sql = "Select Sum (ORDR_SALES) from SOTORDR1 " _
            & " where ORDR_STATUS = 'O' and CUST_CODE = '" & CUST_CODE & "'"
            ORDR_AMT_OPEN = Val(ASCDATA1.GetDataValue)

        End If

        Dim SO_PARM_CRH_MOS_INACTIVE As Int32 = Val(rowSOTPARM1.Item("SO_PARM_CRH_MOS_INACTIVE") & "")
        Dim SO_PARM_CR_LIMIT_GRACE As Int32 = Val(rowSOTPARM1.Item("SO_PARM_CR_LIMIT_GRACE") & "")

        Dim ORDR_HOLD_CREDIT_REASON As String = ""

        ' CRH - Customer Master indicates a Credit Hold
        If rowARTCREDC.Item("CUST_CREDIT_HOLD") & "" = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",CRH"
        End If

        ' SLH - Customer Master indicates a Sales Hold
        Dim CUST_SALES_HOLD As String = rowARTCREDC.Item("CUST_SALES_HOLD") & ""
        If rowARTCREDC.Item("CUST_SALES_HOLD") & "" = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",SLH"
        End If

        ' REV - Customer Master indicates that the Credit Limit Review date is expired
        If rowARTCREDC.Item("CUST_CRED_LIMIT_REV") & "" = "" Then
            'ORDR_HOLD_CREDIT_REASON &= ",REV" ' DON'T HOLD THE ORDER UP IF THERE IS NO REVIEW DATE
        Else
            Dim CUST_CRED_LIMIT_REV As Date = rowARTCREDC.Item("CUST_CRED_LIMIT_REV")
            If Format(CUST_CRED_LIMIT_REV, "yyyyMMdd") _
             < Format(Now + ASCMAIN1.NowTSD, "yyyyMMdd") Then
                ORDR_HOLD_CREDIT_REASON &= ",REV"
            End If
        End If

        If rowARTCREDC.Item("NO_CREDIT_CHECK") & "" <> "1" Then

            ' INA - Last Sale was more than 3 months ago
            If rowARTCREDC.Item("LAST_SALE") & "" = "" Then
                ORDR_HOLD_CREDIT_REASON &= ",INA"
            Else
                Dim LAST_SALE As Date = rowARTCREDC.Item("LAST_SALE")
                If Format(LAST_SALE.AddMonths(SO_PARM_CRH_MOS_INACTIVE), "yyyyMMdd") _
                 < Format(Now + ASCMAIN1.NowTSD, "yyyyMMdd") Then
                    ORDR_HOLD_CREDIT_REASON &= ",INA"
                End If
            End If

            ' LIM - This order puts customer over the Credit Limit, or Customer was already over the Credit Limit
            Dim CUST_CREDIT_LIMIT As Decimal = Val(rowARTCREDC.Item("CUST_CREDIT_LIMIT") & "")
            Dim TOTAL_DUE As Decimal = Val(rowARTCREDC.Item("TOTAL_DUE") & "")
            Dim ORDR_AMT_PT As Decimal = Val(rowARTCREDC.Item("ORDR_AMT_PT") & "")
            Dim ORDR_AMT_PT_REL As Decimal = Val(rowARTCREDC.Item("ORDR_AMT_PT_REL") & "")
            If TOTAL_DUE + ORDR_AMT_PT + ORDR_AMT_OPEN + ORDR_AMT_PT_REL _
             > CUST_CREDIT_LIMIT * (1 + SO_PARM_CR_LIMIT_GRACE / 100) Then
                ORDR_HOLD_CREDIT_REASON &= ",LIM"
            End If

            ' P/D - This Customer has Past Due DR Balances (over 60 days)
            ' P/D - This Customer has Past Due DR Balances (over 90 days) - as per ciro
            Dim PAST_DUE_DR_AMT As Decimal = Val(rowARTCREDC.Item("PAST_DUE_DR_AMT") & "")
            If PAST_DUE_DR_AMT > 0 Then
                ORDR_HOLD_CREDIT_REASON &= ",P/D"
                past_due = True
            End If
        End If

        ' ADD - Customer Master has a COD Add-On Amt Defined, loaded into Order at Order Entry time
        If Val(rowSOTORDR1.Item("ORDR_COD_ADDON_AMT") & "") <> 0 Then
            ORDR_HOLD_CREDIT_REASON &= ",ADD"
        End If

        ' SPC - CSM indicates that order needs special processing
        If rowSOTORDR1.Item("ORDR_HOLD_CREDIT_SPECIAL") & String.Empty = "1" Then
            ORDR_HOLD_CREDIT_REASON &= ",SPC"
        End If

        If TERM_CODE = "C3" And Not credit_card_auth Then
            ' IF TERMS ARE CC AND WE DO NOT HAVE AN AUTH THEN HOLD THE ORDER
            'If ORDR_HOLD_CREDIT_REASON = "" Then
            ORDR_HOLD_CREDIT_REASON &= ",CCA"
            'End If
        Else
            If C1_C2_C3_P9 Or credit_card_auth Then
                If Not past_due And CUST_SALES_HOLD <> "1" Then
                    Return String.Empty
                End If
            End If
        End If

        Return ORDR_HOLD_CREDIT_REASON
    End Function

    Public Shared Sub Initialize_AGING_DATES()
        ReDim Preserve AGING_DATES(4)
        For i As Integer = 0 To 4
            'ascmain1.sql = "Select TO_CHAR(PRD_END_DATE,'dd-MMM-yyyy') from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i)) & "'"
            ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i) & "'"
            AGING_DATES(i) = ASCDATA1.GetDataValue
        Next
    End Sub

    Public Shared Function Calculate_INV_DUE_DATE( _
    ByVal F As ASFBASE1, ByVal TERM_CODE As String, ByVal INV_DATE As Date) As Date

        Dim INV_DUE_DATE As Date = INV_DATE

        If TERM_CODE <> "" Then

            Dim rowTATTERM1 As DataRow = Nothing
            If F IsNot Nothing AndAlso F.dst.Tables.Contains("TATTERM1") Then
                rowTATTERM1 = F.dst.Tables("TATTERM1").Rows.Find(TERM_CODE)
            Else
                rowTATTERM1 = ASCDATA1.GetDataRow("Select * From TATTERM1 Where TERM_CODE = :PARM1", "V", New Object() {TERM_CODE}) '  F.LookUp("TATTERM1", TERM_CODE, True)
            End If
            Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & ""

                Case "D"
                    INV_DUE_DATE = INV_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

                Case "E"

                    Dim ADD_MONTHS_BASE As Integer = 1
                    Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                    Dim BASE_DD As Integer = Val(Format(INV_DATE, "dd"))
                    Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                    Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                    Dim INV_BASE_DATEx As String = Format(INV_DATE, "MM/dd/yyyy")

                    Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & ""
                        Case "F"
                            ASCMAIN1.sql = "Select GLTPARM2.* " _
                             & " from GLTPARM2 " _
                             & " where OPS_YYYYPP = " _
                             & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                             & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_DATE, "dd-MMM-yyyy") & "')"
                            Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow

                            Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                            INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                        Case "C"
                            Dim YYYYMM As String = Format(INV_DATE, "yyyyMM")
                            INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                        Case "S"
                            If BASE_DD <= TERM_CUTOFF_DAY _
                            And BASE_DD <= TERM_DAYS_DUE Then
                                ADD_MONTHS_BASE = 0
                            End If
                            Dim YYYYMM As String = Format(INV_DATE, "yyyyMM")
                            INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                        Case Else
                            INV_DUE_DATE = INV_DATE
                    End Select
                    If TERM_ADDL_MOS > 0 Then
                        INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                    End If

            End Select
        End If

        Return INV_DUE_DATE

    End Function

    Public Shared Sub Record_Event_SOTORDR1( _
    ByVal ORDR_NO As String, _
    ByVal INIT_DATE As Date, _
    ByVal INIT_OPER As String, _
    ByVal EVENT_TYPE As String, _
    ByVal EVENT_DESC As String, _
    Optional ByVal EVENT_KEY As String = "")
        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                             & " Values ('SOTORDR1', :PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6)", _
                             "VDVVVV", _
                             New Object() {ORDR_NO, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY})
    End Sub
    Public Shared Sub Update_ARTCUST6( _
    ByVal frmASFBASE0 As ASFBASE0, _
    ByVal CUST_CODE As String, _
    ByVal rowSOTORDR1 As DataRow, _
    ByVal S As Integer)

        Dim rowARTCUST6 As DataRow = frmASFBASE0.Fill_Record("ARTCUST6", CUST_CODE)
        If rowARTCUST6 Is Nothing Then
            rowARTCUST6 = frmASFBASE0.dst.Tables("ARTCUST6").NewRow
            rowARTCUST6.Item("CUST_CODE") = CUST_CODE
            frmASFBASE0.dst.Tables("ARTCUST6").Rows.Add(rowARTCUST6)
        End If

        With rowARTCUST6
            .Item("CUST_LAST_INV_NUM") = rowSOTORDR1.Item("ORDR_INV_NO")
            .Item("CUST_LAST_INV_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
            .Item("CUST_LAST_INV_AMT") = Val(rowSOTORDR1.Item("ORDR_TOTAL_AMT") & "")
            If .Item("CUST_FIRST_PURCH") & "" = "" Then
                .Item("CUST_FIRST_PURCH") = rowSOTORDR1.Item("ORDR_INV_DATE")
            End If

            If rowSOTORDR1.Item("CUST_CODE") = CUST_CODE Then
                Dim INV_SALES As Decimal = Val(rowSOTORDR1.Item("ORDR_AMT") & "")
                If rowSOTORDR1.Item("ORDR_INV_TYPE") = "I" Then
                    .Item("CUST_SALES_MTD") = Val(.Item("CUST_SALES_MTD") & "") + INV_SALES
                    .Item("CUST_SALES_YTD") = Val(.Item("CUST_SALES_YTD") & "") + INV_SALES
                    .Item("CUST_NUM_INV_MTD") = Val(.Item("CUST_NUM_INV_MTD") & "") + S
                    .Item("CUST_NUM_INV_YTD") = Val(.Item("CUST_NUM_INV_YTD") & "") + S
                ElseIf rowSOTORDR1.Item("ORDR_INV_TYPE") = "C" Then
                    .Item("CUST_CRED_MTD") = Val(.Item("CUST_CRED_MTD") & "") + INV_SALES
                    .Item("CUST_CRED_YTD") = Val(.Item("CUST_CRED_YTD") & "") + INV_SALES
                End If
            End If

            If rowSOTORDR1.Item("CUST_BILL_TO_CUST") = CUST_CODE Then
                ASCMAIN1.sql = "Select Sum (INV_BALANCE) from ARTOPEN1 where CUST_CODE = '" & CUST_CODE & "'"
                Dim CUST_BAL As Decimal = Val(ASCDATA1.GetDataValue)
                If CUST_BAL > Val(.Item("CUST_HIGH_BAL_AMT") & "") Then
                    .Item("CUST_HIGH_BAL_DATE") = rowSOTORDR1.Item("ORDR_INV_DATE")
                    .Item("CUST_HIGH_BAL_AMT") = CUST_BAL
                End If
            End If
        End With
        frmASFBASE0.Update_Record_TDA("ARTCUST6")
    End Sub

    Public Shared Function Get_EDI_Custs(EDI_DOC_NO As String) As List(Of String)

        ASCMAIN1.sql = "Select DISTINCT CUST_CODE from EDTTRPM1 "
        If EDI_DOC_NO <> "" Then
            ASCMAIN1.sql &= " where EDI_DOC_NO = '" & EDI_DOC_NO & "'"
        End If

        Dim c As New List(Of String)
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            c.Add(row.Item(0))
        Next

        Return c
    End Function

    Public Shared Function Prepare_Sales_Invoices( _
        F As ASFBASE1, _
        sqlw As String, _
        ByRef SOTINVH1 As String, _
        ByRef SOTINVH2 As String) As String

        ASCMAIN1.Progress("Building Work File")

        Dim rowGLTPARM2 As DataRow = F.LookUp("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE as ORDR_AMT_SHIP" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST as ORDR_CGS_SHIP" & vbCrLf _
            & " from SOTINVH2, SOTINVH1, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & sqlw

        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE,INV_NO,INV_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_QTY_CANC NUMBER(6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_AMT_CANC NUMBER(13,2)")
        ASCMAIN1.AnalyzeTable(SOTINVH2)


        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SOTORDR1.EDI_APPOINTMENT" & vbCrLf _
            & " from SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and (INV_TYPE, INV_NO) in (Select Distinct INV_TYPE, INV_NO from " & SOTINVH2 & ")"
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE,INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_1 on " & SOTINVH1 & " (INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_2 on " & SOTINVH1 & " (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_3 on " & SOTINVH1 & " (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_4 on " & SOTINVH1 & " (SHIP_BOL_NO)")
        ASCMAIN1.AnalyzeTable(SOTINVH1)

        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ASCMAIN1.sql = "Update " & SOTINVH1 & " Set ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & SOTINVH1 & " Set ORDR_YYYYPP_UPDATED = '" & NYP & "'" _
                & " where INV_DATE > '" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
            ASCDATA1.ExecuteSQL()
        End If


        '& ", 0 TOTAL_UNITS " & vbCrLf _
        '& ", 0 TOTAL_UNITS_CANC " & vbCrLf _
        '& ", 0 TOTAL_UNITS_BACK " & vbCrLf _

        ASCMAIN1.sql = "Select SOTINVH1.* from " & SOTINVH1 & " SOTINVH1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH1", 2))

        ASCMAIN1.sql = "Select SOTINVH2.* from " & SOTINVH2 & " SOTINVH2"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH2", 3))



        ' Credits

        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", NULL ORDR_GROUP_NO, NULL ORDR_ADDR_TYPE_ST, NULL CUST_DC_NO" & vbCrLf _
            & ", NULL EDI_APPOINTMENT" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = 'C'"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHR", 2))

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
            & ", ICTSTYL1.STYLE_COST, ICTSTYL1.SALES_DIVISION_CODE from ICTSTYL1" _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH2 & ")"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))

        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", Decode(E.CUST_CODE,NULL,'N','Y') EDI" & vbCrLf _
            & ", Decode(M.CUST_CODE,NULL,'N','Y') MULTI_STORE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & ", (Select Distinct CUST_CODE from EDTTRPM1 where EDI_STATUS = 'P' and EDI_DOC_NO = '810') E" & vbCrLf _
            & ", (Select CUST_CODE from ARTCUST2 where CUST_ADDR_TYPE = 'MK' group by CUST_CODE having COUNT (*) > 1) M" & vbCrLf _
            & " where E.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and M.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))

        ASCMAIN1.sql = "" _
            & "Select SOTFPCT1.OPS_YYYYPP, SOTFPCT1.CUST_FACTOR_PERCENT, SOTFPCT1.CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTFPCT1" & vbCrLf _
            & " union " & vbCrLf _
            & "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTPARM1 where SO_PARM_KEY = 'Z'" & vbCrLf _
            & " union " & vbCrLf _
            & "Select '" & NYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
            & " from SOTPARM1 where SO_PARM_KEY = 'Z'"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTFPCT1", 1))


        If F.MENU_ITEM_OBJECT = "SORUPDT1" Then
            ASCMAIN1.Progress("-", "Pick Tickets")
            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & vbCrLf _
                & " where SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTINVH1 & ")"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICK1", 1))

            ASCMAIN1.Progress("-", "Shipments")
            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'MK','G:' || SOTSHIP1.ORDR_GROUP_NO, 'S:' || SOTSHIP1.SHIP_BOL_NO) SHIP_BOL_NO_X" & vbCrLf _
                & ", 'N' MULTI_STORE from SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO in (Select DISTINCT SHIP_BOL_NO from " & SOTINVH1 & ")"
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSHIP1", 1))

            ASCMAIN1.Progress("-", "Verify Integrity")
            Dim sql1 As String = "(Select SHIP_BOL_NO, Count (*), Min (INV_NO), Max (INV_NO), Sum (INV_SALES) from " & SOTINVH1 & vbCrLf _
                & " group by SHIP_BOL_NO)" & vbCrLf
            Dim sql2 As String = "(Select SOTSHIP1.SHIP_BOL_NO, Count (*), Min (SOTINVH1.INV_NO), Max (SOTINVH1.INV_NO)" & vbCrLf _
                & ", Sum (SOTINVH1.INV_SALES) from SOTINVH1,SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTINVH1.SHIP_BOL_NO (+) and SOTINVH1.ORDR_YYYYPP_UPDATED IS NULL" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTSHIP1 where REGISTER_XNO is Null)" & vbCrLf _
                & " group by SOTSHIP1.SHIP_BOL_NO)"
            'ASCMAIN1.sql = "" _
            '    & sql1 & " minus " & sql2 & vbCrLf _
            '    & " union " & vbCrLf _
            '    & sql2 & " minus " & sql1
            ASCMAIN1.sql = sql1 & " minus " & sql2
            If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
                Return "Shipments Header Record not in synch w/Invoices; Call ABS"
            End If

            'ASCMAIN1.sql = "Select * from SOTPICK1 where PICK_STATUS <> 'F'"
            'If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
            '    Return "Unconfirmed Pick Tickets linked to Invoices; Call ABS"
            'End If

            If F.dst.Tables("SOTPICK1").Select("ISNULL(PICK_STATUS,'?') <> 'F'").Length <> 0 Then
                Return "Unconfirmed Pick Tickets linked to Invoices; Call ABS"
            End If

            ASCMAIN1.Progress("-", "Pick Ticket Details")
            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTPICK2.PICK_NO in (Select PICK_NO from " & SOTINVH1 & ")" & vbCrLf _
                & " and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & " and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"
            'F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICK2", 2))
            F.Create_TDA(F.dst.Tables.Add, "SOTPICK2", "**", 0, True, "", 2, "PICK_QTY_BACK")
            F.Fill_Records("SOTPICK2")

            F.Create_Relation("SOTINVH1", "SOTPICK2", "PICK_NO")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_CONF)")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS_CANC", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_CANC)")
            F.dst.Tables("SOTINVH1").Columns.Add("TOTAL_UNITS_BACK", GetType(System.Int64), "SUM(CHILD(SOTINVH1_SOTPICK2).PICK_QTY_BACK)")

            F.Create_TDA(F.dst.Tables.Add, "ARTOPEN1", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTCASH1", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTCASH2", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTCASH3", "*")
            F.Create_TDA(F.dst.Tables.Add, "ARTCUST6", "*")
        End If

        ASCMAIN1.Progress("-", "Report Summaries")
        ASCMAIN1.sql = "Select SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE, SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', 'S:' || SOTINVH1.SHIP_BOL_NO,'G:' || SOTORDR1.ORDR_GROUP_NO) AS SHIP_BOL_NO_X" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC','DC','MK') AS SHIP_ADDR_TYPE" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC',SOTORDR1.CUST_DC_NO,'000000') AS SHIP_ADDR_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
            & "  from SOTINVH1," & SOTINVH2 & " SOTINVH2, SOTORDR1" & vbCrLf _
            & "  where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "    and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "    and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & " group by SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE,  SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', 'S:' || SOTINVH1.SHIP_BOL_NO,'G:' || SOTORDR1.ORDR_GROUP_NO)" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC','DC','MK')" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC',SOTORDR1.CUST_DC_NO,'000000')" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf
        Dim SOTINVHD As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_QTY_CANC NUMBER (6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_AMT_CANC NUMBER (13,2)")

        ASCMAIN1.sql = "Update " & SOTINVHD & " Set SHIP_BOL_NO_X = 'G:' || ORDR_GROUP_NO"
        ASCDATA1.ExecuteSQL()

        For Each GROUP As String In New String() {"G"} ' , "S"} NO NEED FOR S SINCE WE MAKE EVERYTHING G ABOVE
            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is Select SOTINVHD.SHIP_BOL_NO_X, SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE" _
                & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" _
                & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_CANC" _
                & "  from SOTORDR1,SOTORDR2," & SOTINVHD & " SOTINVHD" _
                & IIf(GROUP = "S", ",SOTPICK1", "") _
                & "  where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                & "    and SOTORDR2.ORDR_QTY_CANC <> 0" _
                & IIf(GROUP = "G", _
                        "    and SOTORDR1.ORDR_GROUP_NO = SOTINVHD.ORDR_GROUP_NO", _
                        "    and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO and SOTPICK1.SHIP_BOL_NO = SUBSTR(SOTINVHD.SHIP_BOL_NO_X,3)") _
                & "    and SHIP_BOL_NO_X like '" & GROUP & "%'" _
                & "  group by SOTINVHD.SHIP_BOL_NO_X, SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE;" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Begin" _
                & "    Update " & SOTINVHD & " Set " _
                & "      ORDR_QTY_CANC = R1.ORDR_QTY_CANC" _
                & "     ,ORDR_AMT_CANC = R1.ORDR_AMT_CANC" _
                & "     where SALES_DIVISION_CODE = R1.SALES_DIVISION_CODE" _
                & IIf(GROUP = "G", _
                        " and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;", _
                        " and SHIP_BOL_NO_X = R1.SHIP_BOL_NO_X;") _
                & "   End;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()
        Next

        ASCMAIN1.sql = "Select * from " & SOTINVHD
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHD", 0))

        ASCMAIN1.sql = "Select SHIP_BOL_NO_X, SALES_DIVISION_CODE, CUST_CODE, INV_DATE" & vbCrLf _
            & ", Sum (ORDR_QTY_CANC) as QTY_CANC" & vbCrLf _
            & ", Sum (ORDR_AMT_CANC) as AMT_CANC" & vbCrLf _
            & " from " & SOTINVHD & vbCrLf _
            & " group by SHIP_BOL_NO_X, SALES_DIVISION_CODE, CUST_CODE, INV_DATE"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHN", 0))

        ASCMAIN1.sql = "Select SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
            & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
            & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
            & " from " & SOTINVH2 & vbCrLf _
            & " group by SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHC", 0))

        ASCMAIN1.sql = "Select INV_DATE, SALES_DIVISION_CODE, CUST_CODE" & vbCrLf _
            & ", STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
            & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
            & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
            & " from " & SOTINVH2 & vbCrLf _
            & " group by INV_DATE, SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHY", 0))

        ASCMAIN1.sql = "Select * from SOTORDR0 where ORDR_GROUP_NO in" & vbCrLf _
            & " (Select DISTINCT ORDR_GROUP_NO from " & SOTINVH1 & ")"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTORDR0", 0))

        ASCMAIN1.Progress("-", "Consolidated Invoices")

        For Each TABLE_NAME As String In New String() {"SOTINVHZ", "SOTINVHT"}
            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO" & vbCrLf _
                & ", Sum (SOTINVH2.ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, SOTINVH2.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.INV_NO_CONS, SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO" & vbCrLf _
                & ", SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.GST_TAX" & vbCrLf _
                & " from SOTINVH1," & SOTINVH2 & " SOTINVH2" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH1.INV_NO_CONS is " & IIf(TABLE_NAME = "SOTINVHT", "NOT", "") & " Null" & vbCrLf _
                & " group by " & vbCrLf _
                & " SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, SOTINVH2.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.INV_NO_CONS, SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO" & vbCrLf _
                & ", SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.GST_TAX" & vbCrLf
            F.dst.Tables.Add(ASCDATA1.GetDataTable("", TABLE_NAME, 0))
        Next

        For Each rowSOTINVHT As DataRow In F.dst.Tables("SOTINVHT").Select("")
            Dim rowSOTINVHZ As DataRow = F.dst.Tables("SOTINVHZ").NewRow
            rowSOTINVHZ.ItemArray = rowSOTINVHT.ItemArray
            F.dst.Tables("SOTINVHZ").Rows.Add(rowSOTINVHZ)
        Next
        ' HOPEFULLY, THE CODE ABOVE IS GOOD TO REPLACE THE FOLLOWING BLOCKS
        Stop 'HOPE IS NO GOOD
        'For Each row As DataRow In ASCDATA1.SelectDistinct _
        '    (F.dst.Tables("SOTINVHT"), New String() {"INV_NO_CONS", "CUST_CODE", "SALES_DIVISION_CODE", "H_SALES_DIVISION_CODE"}).Select("")
        '    Dim INV_NO_CONS As String = row.Item("INV_NO_CONS")
        '    Dim CUST_CODE As String = row.Item("CUST_CODE")
        '    Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
        '    Dim H_SALES_DIVISION_CODE As String = row.Item("H_SALES_DIVISION_CODE")

        '    Dim sqlw2 As String = "INV_NO_CONS = '{0}' and CUST_CODE = '{1}' and SALES_DIVISION_CODE = '{2}' and H_SALES_DIVISION_CODE = '{3}'"
        '    sqlw2 = String.Format(sqlw2, INV_NO_CONS, CUST_CODE, SALES_DIVISION_CODE, H_SALES_DIVISION_CODE)
        '    Dim rowSOTINVHZ As DataRow = F.dst.Tables("SOTINVHZ").NewRow
        '    With rowSOTINVHZ
        '        .Item("INV_TYPE") = "I"
        '        .Item("INV_NO") = INV_NO_CONS
        '        .Item("INV_DATE") = F.dst.Tables("SOTINVHT").Compute("MAX(INV_DATE)", sqlw2)
        '        .Item("CUST_CODE") = CUST_CODE
        '        .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
        '        .Item("TOTAL_SALES") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(TOTAL_SALES)", sqlw2) & "")
        '        .Item("H_SALES_DIVISION_CODE") = H_SALES_DIVISION_CODE
        '        .Item("ORDR_CUST_PO") = F.dst.Tables("SOTINVHT").Compute("MAX(ORDR_CUST_PO)", sqlw2)
        '        .Item("INV_SALES") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(INV_SALES)", sqlw2) & "")
        '        .Item("INV_FREIGHT") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(INV_FREIGHT)", sqlw2) & "")
        '        .Item("INV_MISC_CHG") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(INV_MISC_CHG)", sqlw2) & "")
        '        .Item("GST_TAX") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(GST_TAX)", sqlw2) & "")
        '        .Item("INV_TOTAL_AMOUNT") = Val(F.dst.Tables("SOTINVHT").Compute("SUM(INV_TOTAL_AMOUNT)", sqlw2) & "")
        '    End With
        '    F.dst.Tables("SOTINVHX").Rows.Add(rowSOTINVHZ)
        'Next

        ''101 01 $100
        ''101 02 $200
        ''Placed Here By Wayne To Resolve Nasty Bug with Consolidated invoices where some of the invoices split divisions and Others don't.
        'Dim dynSOWCONFX As Recordset
        'Sql = "SELECT INV_TYPE, INV_NO, SUM(TOTAL_SALES) AS INV_SALES_NEW"
        '& " FROM SOWINVHZ"
        '& " GROUP BY  INV_TYPE, INV_NO"
        'dynSOWCONFX = AccD.OpenRecordset(Sql, dbOpenForwardOnly)
        'Do While Not dynSOWCONFX.EOF
        '    Dim SQLU As String
        '    SQLU = "Update SOWINVHZ SET INV_SALES = " & dynSOWCONFX.Fields("INV_SALES_NEW").Value
        '    SQLU = SQLU & " WHERE INV_TYPE = '" & dynSOWCONFX.Fields("INV_TYPE").Value & "'"
        '    SQLU = SQLU & " AND INV_NO = '" & dynSOWCONFX.Fields("INV_NO").Value & "'"
        '    AccD.Execute(SQLU)
        '    dynSOWCONFX.MoveNext()
        'Loop
        'dynSOWCONFX.Close()
        ''Finished with Waynes Consolidated Fix.




        'Sql = "SELECT DISTINCT INV_NO, 'X' as UPD"
        '& " Into SOWTEMP"
        '& " From SOWINVHZ"
        'AccD.Execute(Sql)

        'Sql = "Update SOWTEMP, SOWINVHZ set UPD = '1' Where SOWINVHZ.INV_NO = SOWTEMP.INV_NO"
        '& " and SOWINVHZ.SALES_DIVISION_CODE = SOWINVHZ.H_SALES_DIVISION_CODE"
        'AccD.Execute(Sql)

        'Sql = "UPDATE SOWINVHZ, SOWTEMP SET H_SALES_DIVISION_CODE = SALES_DIVISION_CODE "
        '& " WHERE SOWTEMP.UPD = 'X' AND SOWTEMP.INV_NO = SOWINVHZ.INV_NO"
        'AccD.Execute(Sql)



        '& ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO" & vbCrLf _
        '& ", SOTORDR1.EDI_APPOINTMENT" & vbCrLf _
        'sowinvhx goes to xRPT = "SORUPDT1"

        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", 0 AS TOTAL_UNITS" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTINVH1.INV_NO_CONS is Null"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHX", 2))

        ASCMAIN1.sql = "Select SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST as CUST_CODE" & vbCrLf _
            & ", Max(SOTINVH1.INV_DATE) AS INV_DATE" & vbCrLf _
            & ", Max(SOTINVH1.REASON_CODE) as REASON_CODE" & vbCrLf _
            & ", Max(SOTINVH1.SALES_DIVISION_CODE) as SALES_DIVISION_CODE" & vbCrLf _
            & ", Max(SOTINVH1.ORDR_CUST_PO) AS ORDR_CUST_PO" & vbCrLf _
            & ", Max(SOTINVH1.CUST_FACTOR_IND) AS CUST_FACTOR_IND" & vbCrLf _
            & ", Max(SOTINVH1.CUST_SURCHARGE_IND) AS CUST_SURCHARGE_IND" & vbCrLf _
            & ", Sum(SOTINVH1.INV_SALES) AS INV_SALES" & vbCrLf _
            & ", Sum(SOTINVH1.INV_FREIGHT) AS INV_FREIGHT" & vbCrLf _
            & ", Sum(SOTINVH1.INV_MISC_CHG) AS INV_MISC_CHG" & vbCrLf _
            & ", Sum(SOTINVH1.GST_TAX) AS GST_TAX" & vbCrLf _
            & ", Sum(SOTINVH1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT" & vbCrLf _
            & ", 0 AS TOTAL_UNITS" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
            & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_NO_CONS is Not Null" & vbCrLf _
            & " group by SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST" & vbCrLf
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim rowSOTINVHX As DataRow = F.dst.Tables("SOTINVHX").NewRow
            With rowSOTINVHX
                For Each DCOL As DataColumn In row.Table.Columns
                    If DCOL.ColumnName = "INV_NO_CONS" Then
                        .Item("INV_TYPE") = "I"
                        .Item("INV_NO") = row.Item("INV_NO_CONS")
                    Else
                        .Item(DCOL.ColumnName) = row.Item(DCOL.ColumnName)
                    End If
                Next
                .Item("CURR_CODE") = "USD"
                .Item("CURR_EXCH_RATE") = 1
            End With
            F.dst.Tables("SOTINVHX").Rows.Add(rowSOTINVHX)
        Next

        For Each TABLE_NAME As String In New String() {"SOTINVHG1", "SOTINVHG", "SOTINVHG2"}
            With F.dst.Tables.Add(TABLE_NAME)
                .Columns.Add("SD")
                .Columns.Add("CC")
                If TABLE_NAME <> "SOWINVHG2" Then .Columns.Add("ID", GetType(System.DateTime))
                .Columns.Add("QC", GetType(System.Int64))
                .Columns.Add("AC", GetType(System.Decimal))
                If TABLE_NAME = "SOTINVHG" Then .PrimaryKey = New DataColumn() {.Columns("SD"), .Columns("CC"), .Columns("ID")}
                If TABLE_NAME = "SOTINVHG2" Then .PrimaryKey = New DataColumn() {.Columns("SD"), .Columns("CC")}
            End With
        Next

        For Each row As DataRow In F.dst.Tables("SOTINVHN").Select("")
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim INV_DATE As Date = row.Item("INV_DATE")
            F.dst.Tables("SOTINVHG1").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE, row.Item("QTY_CANC"), row.Item("AMT_CANC")})
            If F.dst.Tables("SOTINVHG").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE}) Is Nothing Then
                F.dst.Tables("SOTINVHG").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE})
            End If
            If F.dst.Tables("SOTINVHG2").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE}) Is Nothing Then
                F.dst.Tables("SOTINVHG2").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE})
            End If
        Next

        F.Create_Relation("SOTINVHG", "SOTINVHG1", "SD,CC,ID")
        F.dst.Tables("SOTINVHG").Columns("QC").Expression = "SUM(CHILD.QC)"
        F.dst.Tables("SOTINVHG").Columns("AC").Expression = "SUM(CHILD.AC)"

        F.Create_Relation("SOTINVHG2", "SOTINVHG", "SD,CC")
        F.dst.Tables("SOTINVHG2").Columns("QC").Expression = "SUM(CHILD.QC)"
        F.dst.Tables("SOTINVHG2").Columns("AC").Expression = "SUM(CHILD.AC)"

        ' Master Files

        ASCMAIN1.sql = "Select ARTREAS1.* from ARTREAS1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTREAS1", 1))

        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        F.dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))

        Return ""
    End Function


    Public Shared Sub Track_Shipment(ByVal SHIP_VIA_CODE As String, ByVal SHIP_REF As String)

        ASCMAIN1.ActiveForm.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Locating POD")
        ASCMAIN1.sql = "Select CARRIER_URL_TRACKING, CARRIER_TRACKING_IND" _
        & " from SOTCARR1,SOTSVIA1 " _
        & " where SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'" _
        & "   and SOTCARR1.CARRIER_CODE = SOTSVIA1.CARRIER_CODE"

        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, True)
        Dim CARRIER_URL_TRACKING As String = rowSOTCARR1.Item("CARRIER_URL_TRACKING") & String.Empty
        Dim CARRIER_TRACKING_IND As String = rowSOTCARR1.Item("CARRIER_TRACKING_IND") & String.Empty

        If CARRIER_TRACKING_IND = "I" Then
            ASCMAIN1.sql = "SELECT NVL(INV_NO_RESHIP, INV_NO) FROM SOTINVH1 WHERE SHIP_REF = :PARM1 AND SHIP_VIA_CODE = :PARM2"
            SHIP_REF = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {SHIP_REF, SHIP_VIA_CODE}) & String.Empty
        End If

        If CARRIER_URL_TRACKING = "" Then
            MsgBox("Cannot Determine Carrier Tracking Website URL for Ship Via " & SHIP_VIA_CODE, MsgBoxStyle.OkOnly, "Unable to Perform Requested Action")
        ElseIf SHIP_REF.Length = 0 AndAlso CARRIER_TRACKING_IND = "I" Then
            MsgBox("Cannot Locate tracking information from Tracking Number " & SHIP_REF, MsgBoxStyle.OkOnly, "Unable to Perform Requested Action")
        Else
            System.Diagnostics.Process.Start(CARRIER_URL_TRACKING & SHIP_REF)
        End If
        ASCMAIN1.ActiveForm.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Shared Function Get_EDI_row(EDI_DOC_SEQ_NO As String, _
                                       Optional EDI_DOC_NO As String = "850") As DataRow
        Dim TABLE_NAME As String = "EDT850T1"
        If EDI_DOC_NO = "852" Then TABLE_NAME = "EDT852T1"
        If EDI_DOC_NO = "940" Then TABLE_NAME = "EDT940O1"
        If EDI_DOC_NO = "945" Then TABLE_NAME = "EDT945T1"

        'ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
        '    & " from GEN.^Document_tb^ GD, EDT850T1 " _
        '    & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
        '    & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'" _
        '    & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCMAIN1.sql = "Select GD.*" _
            & " from GEN.^Document_tb^ GD, EDT850T1 " _
            & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
            & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'" _
            & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDT850T1", TABLE_NAME)
        If EDI_DOC_NO = "940" Then
            'ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
            '    & " from GEN.^Document_tb^ GD, EDT940O1 " _
            '    & " where GD.^DocumentName^ like '" & EDI_DOC_SEQ_NO & "%'" _
            '    & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
            ASCMAIN1.sql = "Select GD.* " _
                & " from GEN.^Document_tb^ GD " _
                & " where GD.^DocumentName^ like '" & EDI_DOC_SEQ_NO & "%'" _
                & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'"
        End If
        'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
        '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENAHA.")
        'End If

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "^", Chr(34))

        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.DBS_SERVER = "" Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GEN" & ASCMAIN1.CLIENT & ".")
        End If

        Return ASCDATA1.GetDataRow
    End Function

    Public Shared Function Get_Raw_EDI(EDI_DOC_SEQ_NO As String, _
                                       Optional ED_PARM_RAW_ARCHIVE As String = "", _
                                       Optional EDI_DOC_NO As String = "850") As String

        Dim TABLE_NAME As String = "EDT850T1"
        If EDI_DOC_NO = "852" Then TABLE_NAME = "EDT852T1"
        If EDI_DOC_NO = "820" Then TABLE_NAME = "EDT820T1"
        If EDI_DOC_NO = "180" Then TABLE_NAME = "EDT180T1"
        If EDI_DOC_NO = "812" Then TABLE_NAME = "EDT812T1"
        If EDI_DOC_NO = "850" Then TABLE_NAME = "EDT850T1"
        If EDI_DOC_NO = "830" Then TABLE_NAME = "EDT830T1"

        ASCMAIN1.sql = "Select GD.^DocumentBlobKEY^ RAW_DATA_FILE " _
            & " from GEN.^Document_tb^ GD, EDT850T1 " _
            & " where EDT850T1.GEN_DOC_NO = GD.^AppField1^" _
            & "   and GD.^TransactionSetID^ = '" & EDI_DOC_NO & "'" _
            & "   and EDT850T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDT850T1", TABLE_NAME)

        'If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
        '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENAHA.")
        'End If

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "^", Chr(34))

        '   If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
        '  ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GEN" & ASCMAIN1.CLIENT & ".")
        ' End If

        If ED_PARM_RAW_ARCHIVE = "" Then
            If Not ASCMAIN1.ActiveForm.ROWs.ContainsKey("EDTPARM1") Then
                ASCMAIN1.ActiveForm.Get_PARM("EDTPARM1")
            End If
            ED_PARM_RAW_ARCHIVE = ASCMAIN1.ActiveForm.ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE")

            '   ED_PARM_RAW_ARCHIVE = "\\192.168.175.103\gensrvnt\Documents\"

        End If

        Dim RAW_DATA As String = ""
        Dim RAW_DATA_FILE As String = ASCDATA1.GetDataValue
        If RAW_DATA_FILE <> "" Then
            'Dim FILENAME As String = "V:\Documents\" & RAW_DATA_FILE & ".DOC"
            If ASCMAIN1.Running_in_VS Then
                ED_PARM_RAW_ARCHIVE = "\\Nymain-abs-iis1\gensrvnt\Documents\"
            End If
            Dim FILENAME As String = ED_PARM_RAW_ARCHIVE & RAW_DATA_FILE & ".DOC"
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                RAW_DATA = My.Computer.FileSystem.ReadAllText(FILENAME)
            End If
        End If
        Return RAW_DATA
    End Function

    Public Shared Sub Log_Changes( _
                                frmASFBASE0 As ASFBASE0, _
                                ORDR_NO As String, _
                                row As DataRow, _
                                TABLE_NAME As String, _
                                ByRef Check_Changed_Fields As Boolean, _
                                REV_NO As Integer,
                                ByRef REV_LNO As Integer, _
                                LAST_DATE As Date)

        For i As Integer = 0 To row.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = frmASFBASE0.dst.Tables(TABLE_NAME).Columns(i).ColumnName
            If row.Item(COLUMN_NAME) & "" _
            <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then

                Dim rowSOTORDXR As DataRow = frmASFBASE0.dst.Tables("SOTORDXR").NewRow
                With rowSOTORDXR
                    .Item("REV_NO") = REV_NO
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = row.Item(COLUMN_NAME)
                    .Item("EMODE") = frmASFBASE0.EntryMode
                End With
                frmASFBASE0.dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)

                Check_Changed_Fields = True
            End If
        Next i
    End Sub

    Public Shared Function Get_Price( _
                                    frmASFBASE0 As ASFBASE0, _
                                    PRICE_LIST_CODE As String, _
                                    PRICE_LIST_CODE_ALLO As String, _
                                    PRICE_BASIS As String, _
                                    PRICE_BASE_DPCT As Decimal, _
                                    ITEM_CODE As String, _
                                    rowICTITEM1 As DataRow, _
                                    ORDR_DATE_BOOKED As Date, _
                                    ByRef ITEM_RETAIL_PRICE As Decimal) As Decimal

        Dim ORDR_UNIT_PRICE As Decimal = 0
        ' ITEM_RETAIL_PRICE = 0

        Dim rowSOTPRIC2 As DataRow = Nothing
        If PRICE_LIST_CODE <> "" Then
            rowSOTPRIC2 = frmASFBASE0.LookUp("SOTPRIC2", New String() {PRICE_LIST_CODE, ITEM_CODE})
        End If

        If rowSOTPRIC2 Is Nothing Then
            If PRICE_LIST_CODE_ALLO <> "" Then
                rowSOTPRIC2 = frmASFBASE0.LookUp("SOTPRIC2", New String() {PRICE_LIST_CODE_ALLO, ITEM_CODE})
            End If
        End If

        If rowSOTPRIC2 IsNot Nothing Then
            If rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & "" <> "" _
                AndAlso Format(rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE"), "yyyyMMdd") <= Format(ORDR_DATE_BOOKED, "yyyyMMdd") Then
                ORDR_UNIT_PRICE = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
            Else
                ORDR_UNIT_PRICE = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
            End If
        Else
            If PRICE_BASIS = "R" Then
                'Dim ITEM_RETAIL_PRICE As Decimal
                ITEM_RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                If frmASFBASE0.Name = "SOFRMAF1" Then ' PROB DON'T WANT TO GET OLDER RETAIL PRICE FOR NEW SALES, REGARDLESS OF VALUE OF DATE BOOKED
                    ASCMAIN1.sql = "Select * from ICTRETLA where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                        & " and OPS_YYYYPP = (Select MIN(OPS_YYYYPP) from GLTPARM2" & vbCrLf _
                        & " where PRD_END_DATE >= '" & Format(ORDR_DATE_BOOKED, "dd-MMM-yyyy") & "')"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    If row IsNot Nothing Then
                        ITEM_RETAIL_PRICE = Val(row.Item("ITEM_RETAIL_PRICE") & "")
                    End If
                End If

                If frmASFBASE0.Name = "SOFORDR1" Then
                    Dim ITEM_RETAIL_PRICE_NEW As Decimal = Val(rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE") & "")
                    Dim ITEM_NEW_RETAIL_PRICE_DATE As String = rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE") & ""

                    If ITEM_NEW_RETAIL_PRICE_DATE <> "" And ITEM_RETAIL_PRICE_NEW <> 0 Then
                        If Format(ORDR_DATE_BOOKED, "yyyyMMdd") >= Format(CDate(ITEM_NEW_RETAIL_PRICE_DATE), "yyyyMMdd") Then
                            ITEM_RETAIL_PRICE = ITEM_RETAIL_PRICE_NEW
                        End If
                    End If
                End If

                ORDR_UNIT_PRICE = ITEM_RETAIL_PRICE * (100 - PRICE_BASE_DPCT) / 100
                ORDR_UNIT_PRICE = System.Math.Round(ORDR_UNIT_PRICE, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        Return ORDR_UNIT_PRICE
    End Function

    Public Shared Function SetMonthlyCommissions() As Boolean

        Try

            SetMonthlyCommissions = True

            Dim sql As String = String.Empty

            ' Retail Sales
            sql = " INSERT INTO RSTRETLC  SELECT CUST_CODE, CUST_STORE_NO, OPS_YYYYPP, SUM(QTY_SOLD) QTY_SOLD, SUM(AMT_SOLD) AMT_SOLD"
            sql &= " , NULL, NULL, NULL, NULL, NULL"
            sql &= "  FROM RSTRETL4"
            sql &= " where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            sql &= "  GROUP BY CUST_CODE, CUST_STORE_NO, OPS_YYYYPP"
            ASCDATA1.ExecuteSQL(sql)

            ' Update retail sales with Sell-Thru Rep
            sql = " BEGIN DECLARE CURSOR C1 IS"
            sql &= " SELECT CUST_CODE, CUST_STORE_NO, SELL_CODE "
            sql &= " 	   FROM ARTCUST2 "
            sql &= " 	   WHERE (CUST_CODE, CUST_STORE_NO) IN"
            sql &= " 	   (SELECT DISTINCT CUST_CODE, CUST_STORE_NO FROM RSTRETLC WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "');"
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= "                   UPDATE RSTRETLC "
            sql &= " 				  	SET SELL_CODE = R1.SELL_CODE"
            sql &= " 						 WHERE  OPS_YYYYPP = '" & ASCMAIN1.CYP & "' "
            sql &= " 						 AND CUST_CODE = R1.CUST_CODE "
            sql &= " 						 AND CUST_STORE_NO = R1.CUST_STORE_NO;"
            sql &= " END LOOP;"
            sql &= " END;"
            sql &= " END;"
            ASCDATA1.ExecuteSQL(sql)

            ' Update Sell Thru Comm %
            sql = " UPDATE RSTRETLC SET SELL_COMM_PCT = "
            sql &= " (SELECT NVL(SELL_COMM_PCT, 0) FROM SOTSREP1 WHERE SREP_CODE = RSTRETLC.SELL_CODE)"
            sql &= " WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL(sql)

            ' Set all non assigned Retail Sales to 98 show they show up on the commission report
            sql = " UPDATE RSTRETLC SET SELL_CODE = '98', SELL_COMM_PCT = 0 WHERE SELL_CODE IS NULL AND OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL(sql)

            ' Invoices / Credits - Update Comm % based on current sales rep Comm %
            sql = " Update SOTINVH1 set SREP_COMM_IND = '0'  where ordr_yyyypp_updated = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE SOTINVH1 "
            sql &= " SET SREP_COMM_PCT = (SELECT NVL(SREP_COMM_PCT, 0) FROM SOTSREP1 WHERE SREP_CODE = SOTINVH1.SREP_CODE)"
            sql &= " WHERE SOTINVH1.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'"
            ASCDATA1.ExecuteSQL(sql)

            ' Overrides for Inital sales orders
            sql = "INSERT INTO SOTSCOMO"
            sql &= " SELECT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO"
            sql = sql & "  , SOTINVH1.INV_SALES, SOTINVH1.SREP_CODE, 'Initial Order', 5, NULL, '0', NULL "
            sql = sql & "  FROM SOTINVH1, SOTORDR1 "
            sql = sql & "  WHERE SOTINVH1.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'"
            sql = sql & "  AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO"
            sql = sql & "  AND SOTORDR1.ORDR_INITIAL = '1'"
            sql = sql & "  AND SOTINVH1.INV_SALES <> 0 "
            ASCDATA1.ExecuteSQL(sql)

            ' Invoice Overrides, gets override only if they are not the sales rep on the Invoice
            sql = " INSERT INTO SOTSCOMO"
            sql &= " SELECT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO"
            sql &= " , SOTINVH1.INV_SALES, SOTSREP2.SREP_CODE, 'Override', SOTSREP2.SREP_COMM_PCT, NULL, '0', NULL "
            sql &= " FROM SOTINVH1, SOTSREP2 "
            sql &= " WHERE SOTINVH1.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'"
            sql &= " AND SOTINVH1.INV_SALES <> 0 "
            sql &= " AND SOTSREP2.SREP_COMM_PCT <> 0"
            sql &= " AND SOTINVH1.CUST_CODE = SOTSREP2.CUST_CODE"
            sql &= " AND SOTINVH1.SREP_CODE <> SOTSREP2.SREP_CODE"
            sql &= " AND (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_YYYYPP_UPDATED) NOT IN "
            sql &= " (SELECT INV_TYPE, INV_NO, ORDR_YYYYPP_UPDATED FROM SOTSCOMO)"
            ASCDATA1.ExecuteSQL(sql)

        Catch ex As Exception
            SetMonthlyCommissions = False
        End Try

    End Function

    Public Shared Function Validate_Invoice_Date(DT As Date, MOS_BACK As Integer, MOS_FWD As Integer, ByRef EMsg As String) As Date()
        ASCMAIN1.sql = String.Format("SELECT DTE1, DTE2 FROM " _
            & "(SELECT PRD_END_DATE+1 DTE1 FROM GLTPARM2 WHERE OPS_YYYYPP = {0})," _
            & "(SELECT PRD_END_DATE DTE2 FROM GLTPARM2 WHERE OPS_YYYYPP = {1})", _
            ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 - MOS_BACK), ASCMAIN1.Period_Calc(ASCMAIN1.CYP, MOS_FWD))

        Dim row As DataRow = ASCDATA1.GetDataRow
        Dim dte1 As Date = row.Item("DTE1")
        Dim dte2 As Date = row.Item("DTE2")

        If DT & "" <> "" Then
            If Format(DT, "yyyyMMdd") < Format(dte1, "yyyyMMdd") _
            Or Format(DT, "yyyyMMdd") > Format(dte2, "yyyyMMdd") Then
                EMsg &= vbCr & "Valid Date Range is " & Format(dte1, "MM/dd/yyyy") & " thru " & Format(dte2, "MM/dd/yyyy")
            End If
        End If

        Return New Date() {dte1, dte2}
    End Function

    Public Shared Function IssueCredit(ByVal INV_NO As String, ByRef ErrorMessage As String) As Boolean

        MsgBox("Method not Supported (Need Credit Card Integrator")
        Return False

        ' the code block below was remmed out by WJZ on 06/17 to start implementation of ACH for INT using ePayment Integrator

        'Try

        '    IssueCredit = False

        '    Dim CCPA_NO As String = String.Empty

        '    ' Only process in Live Environment
        '    If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.DBS_SERVER Then
        '        Return True
        '    End If

        '    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_NO = :PARM1", "V", INV_NO)

        '    If rowSOTINVH1 Is Nothing Then
        '        ErrorMessage = "Cannot Locate the supplied Credit Invoice No: " & INV_NO
        '        Exit Function
        '    End If

        '    If rowSOTINVH1.Item("CC_CRED_TRANS_ID") & String.Empty <> String.Empty Then
        '        ErrorMessage = "The credit was already issued a refund using a credit card."
        '        Exit Function
        '    End If

        '    Dim CreditAmount As Decimal = Math.Abs(Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty))
        '    If CreditAmount = 0 Then
        '        ErrorMessage = "Cannot process Credit for $0.00 credit"
        '        Exit Function
        '    End If

        '    Dim Transaction_ID As String = (rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty).ToString.Trim
        '    If Transaction_ID.Length = 0 Then
        '        ErrorMessage = "Invalid or Missing Transaction ID"
        '        Exit Function
        '    End If

        '    Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE TRANS_ID = :PARM1", "V", New Object() {Transaction_ID})
        '    If rowARTCCPA1 Is Nothing Then
        '        ErrorMessage = "Invalid or Missing Transaction ID"
        '        Exit Function
        '    End If

        '    Dim creditCard As String = (rowARTCCPA1.Item("CUST_CREDIT_CARD_LAST4") & String.Empty).ToString.Trim
        '    If creditCard.Length <> 4 Then
        '        creditCard = (rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty).ToString.Trim
        '    End If

        '    ' Get the last 4
        '    If creditCard.Length > 0 Then
        '        creditCard = StrReverse(StrReverse(creditCard).Substring(0, 4))
        '    End If

        '    If creditCard.Length <> 4 Then
        '        ErrorMessage = "Invalid or Missing Credit Card Number"
        '        Exit Function
        '    End If

        '    Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = :PARM1", "V", "Z")
        '    Dim ff As New ABSolution.ASFBASE1
        '    Dim CreditCardProcessor As New TAC.TAFCARDF(ff)

        '    Try
        '        CreditCardProcessor.test_mode = rowSOTPARM1.Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
        '        CreditCardProcessor.CUST_CODE = rowSOTINVH1.Item("CUST_CODE") & String.Empty
        '        CreditCardProcessor.CCPA_REASON = "M"
        '        CreditCardProcessor.TRAN_TYPE = "C"
        '        CreditCardProcessor.MerchantSetup()
        '        CreditCardProcessor.rowARTCCPA1 = rowARTCCPA1
        '        CCPA_NO = CreditCardProcessor.CC_Credit(Transaction_ID, CreditAmount, creditCard)
        '    Catch ex As Exception
        '        ErrorMessage = "Error Processing Credit Card Refund: " & ex.Message
        '        Return False
        '    End Try

        '    If CCPA_NO.Length > 0 Then
        '        ASCMAIN1.sql = "Update SOTINVH1 set CCPA_NO = '" & CCPA_NO & "', CC_CRED_TRANS_ID = '" & CreditCardProcessor.MerchantTransID & "' WHERE INV_TYPE = 'C' AND INV_NO = '" & INV_NO & "'"
        '        ASCDATA1.ExecuteSQL()

        '        ASCMAIN1.sql = "Update ARTCCPA1 set INV_NO = '" & INV_NO & "' WHERE CCPA_NO = '" & CCPA_NO & "'"
        '        ASCDATA1.ExecuteSQL()

        '        ErrorMessage = CreditCardProcessor.MerchantTransID
        '        IssueCredit = True
        '    Else
        '        ErrorMessage = "Could not process Credit Card Refund for the following reason: " & CreditCardProcessor.responseErrorMessage
        '        Return False
        '    End If

        'Catch ex As Exception
        '    ErrorMessage = "Error Processing Credit Card Refund: " & ex.Message
        '    Return False
        'End Try

    End Function

    Public Shared Sub GetCreditCardSaleTransaction(ByVal CustomerCode As String, _
                                            ByVal ReferenceNo As String, _
                                            ByRef OriginalInvNo As String, _
                                            ByRef CreditCardSaleTransaction As String)

        Try


            Dim ccTransFound As Boolean = False
            ReferenceNo = ReferenceNo.Trim
            CustomerCode = CustomerCode.Trim

            If ReferenceNo.Length = 0 OrElse CustomerCode.Length = 0 Then
                Exit Sub
            End If

            ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE INV_TYPE = 'I' AND CUST_CODE = :PARM1 AND ORDR_CUST_PO = :PARM2 AND CC_SALE_TRANS_ID IS NOT NULL"
            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {CustomerCode, ReferenceNo})
            If rowSOTINVH1 IsNot Nothing AndAlso rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty Then
                CreditCardSaleTransaction = rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty
                OriginalInvNo = rowSOTINVH1.Item("INV_NO") & String.Empty
                Exit Sub
            End If

            ' Try to get a hit by Invoice Number
            Dim INV_NO As String = ASCMAIN1.Format_Field(ReferenceNo, "INV_NO")
            ASCMAIN1.sql = "SELECT * FROM SOTINVH1 WHERE INV_TYPE = 'I' AND CUST_CODE = :PARM1 AND INV_NO = :PARM2 AND CC_SALE_TRANS_ID IS NOT NULL"
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {CustomerCode, INV_NO})
            If rowSOTINVH1 IsNot Nothing AndAlso rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty Then
                CreditCardSaleTransaction = rowSOTINVH1.Item("CC_SALE_TRANS_ID")
                OriginalInvNo = rowSOTINVH1.Item("INV_NO")
                Exit Sub
            End If

            ' Try to get a hot by order Number and use the CCPA NO
            Dim ORDR_NO As String = ASCMAIN1.Format_Field(ReferenceNo, "ORDR_NO")
            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE CUST_CODE = :PARM1 AND ORDR_NO = :PARM2"
            rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {CustomerCode, ORDR_NO})
            If rowSOTINVH1 IsNot Nothing AndAlso rowSOTINVH1.Item("CC_TRANS_ID") & String.Empty <> String.Empty Then
                CreditCardSaleTransaction = rowSOTINVH1.Item("CC_TRANS_ID")
                OriginalInvNo = ASCDATA1.GetDataValue("SELECT INV_NO FROM SOTINVH1 WHERE ORDR_NO = '" & ORDR_NO & "' AND INV_TYPE = 'I'") & String.Empty
            ElseIf rowSOTINVH1 IsNot Nothing AndAlso rowSOTINVH1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                Dim CCPA_NO As String = rowSOTINVH1.Item("CCPA_NO") & String.Empty
                ASCMAIN1.sql = "SELECT * FROM ARTCCPA1 WHERE CCPA_NO = :PARM1"
                rowSOTINVH1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {CCPA_NO})
                If rowSOTINVH1 IsNot Nothing AndAlso rowSOTINVH1.Item("TRANS_ID") & String.Empty <> String.Empty Then
                    CreditCardSaleTransaction = rowSOTINVH1.Item("TRANS_ID")
                    OriginalInvNo = ASCDATA1.GetDataValue("SELECT INV_NO FROM SOTINVH1 WHERE ORDR_NO = '" & ORDR_NO & "' AND INV_TYPE = 'I'") & String.Empty
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Shared Sub Build_ARTCSUMC( _
        frmASFBASE0 As ASFBASE0, _
        CUST_CODE As String, _
        YP0 As String, _
        YP1 As String, _
        ARTCSUMC As String, _
        Optional clear_ARTCSUMB As Boolean = True)

        If clear_ARTCSUMB Then
            ASCMAIN1.sql = "Truncate Table " & ARTCSUMC
            ASCDATA1.ExecuteSQL()
        End If

        Dim Xs As String = ""
        For I As Integer = 1 To 12
            Xs &= ", X.AMT" & Format(I, "00")
        Next

        '        Dim sqls As New Dictionary(Of String, String)

        Dim sqlSOTINVH2where As String = "" _
            & " where SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & YP0 & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & YP1 & "'" & vbCrLf

        Dim sqlARTPYMT5where As String = "" _
            & " where ARTPYMT2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & YP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & YP1 & "'" & vbCrLf _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf

        Dim sqlARTPYMT4where As String = "" _
            & " where ARTPYMT2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & YP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & YP1 & "'" & vbCrLf _
            & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf

        Dim sqlSOTINVH1where As String = "" _
            & " where SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & YP0 & "'" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & YP1 & "'" & vbCrLf _
            & "   and SOTINVH1.REASON_CODE is Not Null" & vbCrLf _
            & "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf

        Dim SHP As String = ""
        Dim DED As String = ""
        Dim CRM As String = ""
        Dim GLD As String = ""
        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(YP0, I - 1)
            SHP &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) AMT" & Format(I, "00")
            DED &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & YP & "',DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'0',ARTPYMT5.GL_DIST_AMT,0),0)) AMT" & Format(I, "00")
            CRM &= ", SUM (DECODE(SOTINVH1.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH1.INV_TOTAL_AMOUNT,0),0)) AMT" & Format(I, "00")
            GLD &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & YP & "',ARTPYMT4.GL_DIST_AMT,0)) AMT" & Format(I, "00")
        Next


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 10 LINE, X.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTITEM1, (Select SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & SHP & vbCrLf _
            & " from SOTINVH2" & vbCrLf _
            & sqlSOTINVH2where _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        ASCDATA1.ExecuteSQL()

        'Dim LINE As String = "DECODE(SOTINVH1.ORDR_TYPE_CODE,'DIF',21,20)"
        'ASCMAIN1.sql = "Insert into " & ARTCSUMC _
        '    & " Select X.LINE, X.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
        '    & Xs & vbCrLf _
        '    & " from ICTITEM1, (Select " & LINE & " LINE, SOTINVH2.ITEM_CODE" & vbCrLf _
        '    & SHP & vbCrLf _
        '    & " from SOTINVH2,SOTINVH1" & vbCrLf _
        '    & sqlSOTINVH2where _
        '    & "   and SOTINVH2.INV_TYPE = 'C'" & vbCrLf _
        '    & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
        '    & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
        '    & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
        '    & " group by " & LINE & ",SOTINVH2.ITEM_CODE) X" & vbCrLf _
        '    & " where ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 20 LINE, X.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTITEM1, (Select SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & SHP & vbCrLf _
            & " from SOTINVH2,SOTINVH1" & vbCrLf _
            & sqlSOTINVH2where _
            & "   and SOTINVH2.INV_TYPE = 'C'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE > 0" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTITEM1.ITEM_SNU_CODE,'S',40" & vbCrLf _
            & ", DECODE(ICTPROD1.PROD_CATGY,'D',64" & vbCrLf _
            & ", DECODE(ICTPROD1.PROD_CATGY,'G',61,60))) LINE" & vbCrLf _
            & ", X.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ICTITEM1, ICTCOST1, ICTPROD1" & vbCrLf _
            & ", (Select SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & Replace(SHP, "ORDR_UNIT_PRICE", "ITEM_UNIT_COST") & vbCrLf _
            & " from SOTINVH2" & vbCrLf _
            & sqlSOTINVH2where _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.ITEM_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE (+) = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 80 LINE, X.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ARTREAS1, (Select ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & DED & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
            & sqlARTPYMT5where _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE) X" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = X.REASON_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, DECODE(X.INV_TYPE,'I',71,70) LINE, X.REASON_CODE, ARTREAS1.REASON_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from ARTREAS1, (Select SOTINVH1.CUST_CODE, SOTINVH1.INV_TYPE, SOTINVH1.REASON_CODE" & vbCrLf _
            & CRM & vbCrLf _
            & " from SOTINVH1" & vbCrLf _
            & sqlSOTINVH1where _
            & " group by SOTINVH1.CUST_CODE, SOTINVH1.INV_TYPE, SOTINVH1.REASON_CODE) X" & vbCrLf _
            & " where ARTREAS1.REASON_CODE (+) = X.REASON_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Insert into " & ARTCSUMC _
            & " Select X.CUST_CODE, 81 LINE, X.ACCT_CODE, GLTACCT1.ACCT_DESC" & vbCrLf _
            & Xs & vbCrLf _
            & " from GLTACCT1, (Select ARTPYMT2.CUST_CODE, ARTPYMT4.ACCT_CODE" & vbCrLf _
            & GLD & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCrLf _
            & sqlARTPYMT4where _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT4.ACCT_CODE) X" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+) = X.ACCT_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & ARTCSUMC & " where CUST_CODE = '" & CUST_CODE & "'"
        frmASFBASE0.Fill_Records("ARTCSUMB", "", clear_ARTCSUMB, ASCMAIN1.sql)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 30, 10, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 30, 20, 1)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 50, 30, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 50, 40, -1)

        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 50, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 60, -1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 61, -1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 64, -1)
        ' intentionally leaving out PPP 65 because we get re-imbursed for it = 
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 70, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 71, 1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 80, -1)
        Build_ARTCSUMA_Totals(frmASFBASE0, CUST_CODE, 99, 81, -1)
    End Sub

    Public Shared Sub Build_ARTCSUMA_Totals( _
        frmASFBASE0 As ASFBASE0, _
        CUST_CODE As String, _
        LINE As Integer, LINE_to_add As Integer, S As Integer)

        Dim rowARTCSUMA As DataRow = frmASFBASE0.dst.Tables("ARTCSUMA").Rows.Find(New Object() {CUST_CODE, LINE_to_add})
        Dim rowARTCSUMB As DataRow = frmASFBASE0.dst.Tables("ARTCSUMB").NewRow
        rowARTCSUMB.Item("CUST_CODE") = CUST_CODE
        rowARTCSUMB.Item("LINE") = LINE
        rowARTCSUMB.Item("CODE_VALUE") = rowARTCSUMA.Item("LINE_ABBR")
        rowARTCSUMB.Item("DESC_VALUE") = rowARTCSUMA.Item("LINE_DESC")

        For I As Integer = 1 To 12
            Dim C As String = "AMT" & Format(I, "00")
            rowARTCSUMB.Item(C) = Val(rowARTCSUMA.Item(C) & "") * S
        Next
        frmASFBASE0.dst.Tables("ARTCSUMB").Rows.Add(rowARTCSUMB)
    End Sub

    Public Shared Sub Create_ARTCSUMA( _
        frmASFBASE0 As ASFBASE0, _
        CUST_CODE As String)

        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 10, "Gross Shipments", "GRS")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 20, "Returns", "RTN")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 30, "Net Sales", "NET")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 40, "Cost of Goods", "CGS")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 50, "$GP on Net Sales", "GP")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 60, "Testers/Samples/Misc", "ST")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 61, "Gift w/Purchase", "GWP")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 64, "Displays", "DSP")
        ' Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 65, "Pre-Paid Promo", "PP")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 70, "Misc Credits", "CR")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 71, "Misc Charges", "DR")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 80, "Deductions", "DED")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 81, "GL Write-Offs", "GL")
        Add_ARTCSUMA(frmASFBASE0, CUST_CODE, 99, "Net Profit", "NP")
    End Sub

    Public Shared Sub Add_ARTCSUMA( _
        frmASFBASE0 As ASFBASE0, _
        CUST_CODE As String, _
        LINE As Integer, LINE_DESC As String, LINE_ABBR As String)
        Dim rowARTCSUMA As DataRow = frmASFBASE0.dst.Tables("ARTCSUMA").NewRow
        rowARTCSUMA.Item("CUST_CODE") = CUST_CODE
        rowARTCSUMA.Item("LINE") = LINE
        rowARTCSUMA.Item("LINE_DESC") = LINE_DESC
        rowARTCSUMA.Item("LINE_ABBR") = LINE_ABBR
        frmASFBASE0.dst.Tables("ARTCSUMA").Rows.Add(rowARTCSUMA)
    End Sub

    Public Shared Function Create_ARTCSUMC(frmASFBASE0 As ASFBASE0) As String

        Dim ARTCSUMC As String = ""

        With frmASFBASE0.dst
            If .Relations.Contains("ARTCSUMA_ARTCSUMB") Then
                For I As Integer = 0 To 13
                    Dim C As String = "AMT" & Format(I, "00")
                    .Tables("ARTCSUMA").Columns.Remove(C)
                Next
                .Relations.Remove("ARTCSUMA_ARTCSUMB")
            End If

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, 0 LINE, ARTCUST1.CUST_NAME CODE_VALUE, ARTCUST1.CUST_NAME DESC_VALUE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ",0.01 AMT" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from ARTCUST1 where ROWNUM < 1"
            ARTCSUMC = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Alter Table " & ARTCSUMC & " Modify DESC_VALUE VARCHAR2(300)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select CUST_CODE, LINE"
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", AMT" & Format(I, "00")
            Next
            ASCMAIN1.sql &= " from " & ARTCSUMC & " ARTCSUMC where CUST_CODE = :PARM1 and LINE = :PARM2"
            frmASFBASE0.Create_TDA(.Tables.Add, "ARTCSUMA", "**", 0, False, "VN", 0)
            With .Tables("ARTCSUMA")
                .Columns.Add("LINE_DESC")
                .Columns.Add("LINE_ABBR")
                .Columns.Add("AMT00", GetType(System.Decimal))
                .Columns.Add("AMT13", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("LINE")}
            End With

            ASCMAIN1.sql = "Select CUST_CODE, LINE, CODE_VALUE, DESC_VALUE"
            Dim T As String = ""
            For I As Integer = 1 To 12
                ASCMAIN1.sql &= ", AMT" & Format(I, "00")
                T &= "+ISNULL(AMT" & Format(I, "00") & ",0)"
            Next
            ASCMAIN1.sql &= " from " & ARTCSUMC & " ARTCSUMC where CUST_CODE = :PARM1 and LINE = :PARM2"
            frmASFBASE0.Create_TDA(.Tables.Add, "ARTCSUMB", "**", 0, False, "VN", 0)
            .Tables("ARTCSUMB").Columns.Add("AMT00", GetType(System.Decimal), Mid(T, 2))
            .Tables("ARTCSUMB").Columns.Add("AMT13", GetType(System.Decimal), Mid(T, 2))

            frmASFBASE0.Create_Relation("ARTCSUMA", "ARTCSUMB", "CUST_CODE,LINE")
            For I As Integer = 0 To 13
                Dim C As String = "AMT" & Format(I, "00")
                .Tables("ARTCSUMA").Columns(C).Expression = "SUM(CHILD(ARTCSUMA_ARTCSUMB)." & C & ")"
            Next
        End With

        Return ARTCSUMC
    End Function

    Public Shared Function Setup_Budgets_by_Customer() As String
        Dim SATBUDW1 As String = "SATBUDW1"

        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then

            ASCMAIN1.sql = "Select * from SATBUDW1"
            SATBUDW1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Create Index I_" & SATBUDW1 & "_I on " & SATBUDW1 & " (CUST_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & SATBUDW1 & " Add Primary Key (OPS_YYYY, BRAND_CODE, CUST_CODE)")

            ASCMAIN1.sql = "Select Distinct CUST_CODE from " & SATBUDW1 & " where CUST_CODE LIKE '%:%'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CUST_CODE_cls As String = row.Item("CUST_CODE")
                If CUST_CODE_cls <> "" Then
                    Dim TRADE_CLASS_CODE As String = Split(CUST_CODE_cls, ":")(0)
                    Dim CUST_CLASS_CODE As String = Split(CUST_CODE_cls, ":")(1)
                    ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1 where TRADE_CLASS_CODE = :PARM1 and CUST_CLASS_CODE = :PARM2 and CUST_CODE Not in (Select DISTINCT CUST_CODE from " & SATBUDW1 & " where CUST_CODE NOT LIKE '%:%')"
                    Dim CUST_CODE As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {TRADE_CLASS_CODE, CUST_CLASS_CODE})
                    If CUST_CODE <> "" Then
                        ASCMAIN1.sql = "Update " & SATBUDW1 & " Set CUST_CODE = :PARM1 where CUST_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {CUST_CODE, CUST_CODE_cls})
                    End If
                End If
            Next
        End If

        Return SATBUDW1
    End Function

    ''' <summary>
    ''' Validates High Collection Authorization
    ''' </summary>
    ''' <param name="tblSOTORDR2">Order Details</param>
    ''' <returns>Returns Error Message for Customer/Store/High Colection not Authorized</returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateAuthorizations(ByVal tblSOTORDR2 As DataTable) As String
        Dim salesOrderList As New List(Of String)
        Return ValidateAuthorizations(tblSOTORDR2, salesOrderList)
    End Function

    ''' <summary>
    ''' Validates High Collection Authorization
    ''' </summary>
    ''' <param name="tblSOTORDR2">Order Details</param>
    ''' <param name="salesOrderList">List to contain order numbers with Authorization problems</param>
    ''' <returns>Returns Error Message for Customer/Store/High Colection not Authorized</returns>
    ''' <remarks></remarks>
    Public Shared Function ValidateAuthorizations(ByVal tblSOTORDR2 As DataTable, ByRef salesOrderList As List(Of String)) As String

        ValidateAuthorizations = String.Empty
        If salesOrderList Is Nothing Then
            salesOrderList = New List(Of String)
        End If

        Dim tblCust As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"CUST_CODE", "CUST_STORE_NO"})
        Dim custStore As String = String.Empty
        Dim custCodes As New List(Of String)

        For Each row As DataRow In tblCust.Select("")
            custStore &= ", ('" & row.Item("CUST_CODE") & "', '" & row.Item("CUST_STORE_NO") & "')"
            If Not custCodes.Contains(row.Item("CUST_CODE")) Then
                custCodes.Add(row.Item("CUST_CODE"))
            End If
        Next

        If custStore.Length = 0 Then
            ValidateAuthorizations = "Authorizations could not determine the Customer / Store No"
            Exit Function
        Else
            custStore = custStore.Substring(1).Trim
            custStore = "( " & custStore & " )"
        End If

        Dim tblItems As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"ITEM_CODE"})
        Dim items As String = String.Empty
        For Each row As DataRow In tblItems.Select("")
            items &= ", '" & row.Item("ITEM_CODE") & "'"
        Next
        If items.Length = 0 Then
            ValidateAuthorizations = "Authorizations could not determine the Items"
            Exit Function
        Else
            items = items.Substring(1).Trim
            items = "( " & items & " )"
        End If

        Dim nextPeriod As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTPARM1 where SO_PARM_KEY = 'Z'")
        Dim SO_PARM_AUTH_MOS_CLOSE As Integer = Val(rowSOTPARM1.Item("SO_PARM_AUTH_MOS_CLOSE") & "")
        Dim YPX As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, SO_PARM_AUTH_MOS_CLOSE)
        '& " AND SATAUTH1.OPS_YYYYPP_OPENED <= '" & ASCMAIN1.CYP & "'" _

        ASCMAIN1.sql = " Select SATAUTH1.*,ICTITEM1.ITEM_CODE" _
            & " FROM SATAUTH1, ICTCOLL1, ICTITEM1" _
            & " WHERE SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE" _
            & " AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
            & " AND (SATAUTH1.CUST_CODE, SATAUTH1.CUST_STORE_NO) in " & custStore _
            & " AND SATAUTH1.OPS_YYYYPP_OPENED <= '" & YPX & "'" _
            & " AND NVL(SATAUTH1.OPS_YYYYPP_CLOSED, '" & nextPeriod & "') > '" & ASCMAIN1.CYP & "'" _
            & " AND ICTITEM1.ITEM_CODE IN " & items
        Dim tblSATAUTH1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT DISTINCT ICTITEM1.ITEM_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_CODE" _
            & " FROM ICTITEM1, ICTCOLL1" _
            & " WHERE ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" _
            & " AND ICTITEM1.ITEM_CODE IN " & items
        Dim tblHC_CODES As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        tblHC_CODES.PrimaryKey = New DataColumn() {tblHC_CODES.Columns("ITEM_CODE")}


        ASCMAIN1.sql = "SELECT COLLECTION_CODE,STATE_CODE,CUST_CODE FROM SATEXCL1"
        Dim tblExclusions As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        Dim ORDR_NO = If(tblSOTORDR2.Rows.Count > 0,tblSOTORDR2.Rows(0).Item("ORDR_NO"),"")
        

        Dim errorList As New List(Of String)
        Dim errorMsg As String = String.Empty

        'Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE")
        'Dim rowSOTTCLS1 As DataRow = Lookup("SOTTCLS1", TRADE_CLASS_CODE)

        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE " _
            & " FROM ARTCUST1, SOTTCLS1" _
            & " WHERE ARTCUST1.TRADE_CLASS_CODE = SOTTCLS1.TRADE_CLASS_CODE" _
            & " AND AUTH_REQD = '1'" _
            & " AND CUST_CODE IN ('" & String.Join("', '", custCodes.ToArray) & "')"

        Dim tblARTCUST1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowARTCUST1 As DataRow In tblARTCUST1.Select("", "CUST_CODE")
            Dim CUST_CODEx As String = rowARTCUST1.Item("CUST_CODE")
            Dim prevStore As String = ""
            Dim stateCode As String = ""
            ' Look only at items with Order Qty Open > 0
            For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("CUST_CODE = '" & CUST_CODEx & "' and ORDR_QTY_OPEN > 0", "CUST_STORE_NO,ITEM_CODE")
                Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE") & String.Empty
                Dim CUST_STORE_NO As String = rowSOTORDR2.Item("CUST_STORE_NO") & String.Empty
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty

                If CUST_STORE_NO <> prevStore THen
                    stateCode = ASCDATA1.GetDataValue($"SELECT CUST_STORE_STATE FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2 ","VV",{CUST_CODE,CUST_STORE_NO}) & ""
                End If
                prevStore = CUST_STORE_NO


                ASCMAIN1.sql = "CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "' AND ITEM_CODE = '" & ITEM_CODE & "'"
                Dim row As DataRow = tblHC_CODES.Rows.Find(ITEM_CODE)

                If tblSATAUTH1.Select(ASCMAIN1.sql, "").Length = 0 Then
                    

                    If row IsNot Nothing Then
                        errorMsg = CUST_CODE & "/" & CUST_STORE_NO & " not Authorized for High Collection: " & row.Item("HC_CODE")
                    Else
                        errorMsg = CUST_CODE & "/" & CUST_STORE_NO & " not Authorized for Item: " & ITEM_CODE
                    End If

                    If Not errorList.Contains(errorMsg) Then
                        errorList.Add(errorMsg)
                    End If

                    If Not salesOrderList.Contains(rowSOTORDR2.Item("ORDR_NO")) Then
                        salesOrderList.Add(rowSOTORDR2.Item("ORDR_NO"))
                    End If
                End If
                
                If tblExclusions.Rows.Find({row.Item("COLLECTION_CODE"),stateCode,CUST_CODE}) IsNot Nothing Then
                    errorMsg = $"Item {ITEM_CODE} (Coll. {row.Item("COLLECTION_CODE")}) excluded for {CUST_CODE}/{CUST_STORE_NO} in {stateCode}"
                    If Not errorList.Contains(errorMsg) Then
                        errorList.Add(errorMsg)
                    End If
                    If Not salesOrderList.Contains(rowSOTORDR2.Item("ORDR_NO")) Then
                        salesOrderList.Add(rowSOTORDR2.Item("ORDR_NO"))
                    End If
                End If
            Next
        Next

        If errorList.Count > 0 Then
            ValidateAuthorizations = String.Join(vbCrLf, errorList.ToArray)
        End If

    End Function

    Public Shared Function ValidateOrderQtys(ByVal tblSOTORDR2 As DataTable _
                                             , Optional FieldName As String = "ORDR_QTY" _
                                             , Optional multiStore As Boolean = False) As String
        Dim salesOrderList As New List(Of String)
        Return ValidateOrderQtys(tblSOTORDR2, salesOrderList, FieldName, multiStore)
    End Function

    Public Shared Function ValidateOrderQtys(ByVal tblSOTORDR2 As DataTable _
                                             , ByRef salesOrderList As List(Of String) _
                                             , Optional FieldName As String = "ORDR_QTY" _
                                             , Optional multiStore As Boolean = False) As String

        ValidateOrderQtys = String.Empty

        Dim tblItems As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"ITEM_CODE"})
        Dim items As String = String.Empty
        For Each row As DataRow In tblItems.Select("")
            items &= ", '" & row.Item("ITEM_CODE") & "'"
        Next

        If items.Length > 0 Then
            items = items.Substring(1).Trim
            items = "( " & items & " )"
        End If

        ASCMAIN1.sql = "Select * from ICTITEM1 WHERE ITEM_CODE IN  " & items
        Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTITEM1")

        Dim errorList As New List(Of String)
        Dim errorMsg As String = String.Empty

        For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "ITEM_CODE")
            Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE") & String.Empty
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty
            Dim rowICTITEM1 As DataRow = tblICTITEM1.Rows.Find(ITEM_CODE)

            If rowICTITEM1 Is Nothing Then
                Continue For
            End If

            Dim PROD_CODE As String = rowICTITEM1.Item("PROD_CODE") & String.Empty
            Dim is_NCS As Boolean = (PROD_CODE = "NCS")

            If CUST_CODE.StartsWith("IPLB") AndAlso Not CUST_CODE.StartsWith("IPLBAE") AndAlso Not is_NCS Then
                ' LM SAYS TO AVOID QTY CHECKS FOR ALL CUSTOMERS BEGINNING WITH IPLB
                'ISSUE-7310: STILL ENFORCE FOR NCS (SAMPLES)
                Continue For
            End If

            Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO") & String.Empty
            Dim ORDR_LNO As String = rowSOTORDR2.Item("ORDR_LNO")
            Dim CUST_STORE_NO As String = rowSOTORDR2.Item("CUST_STORE_NO") & String.Empty

            Dim ORDR_QTY As Int32 = Val(rowSOTORDR2.Item(FieldName) & String.Empty)
            If ORDR_QTY = 0 Then
                Continue For
            End If

            ' SO Min Qty
            Dim ITEM_SO_QTY_MIN As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MIN") & String.Empty)
            ' SO Multiple
            Dim ITEM_SO_QTY_MULT As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & String.Empty)
            ' Inner Pack
            Dim ITEM_STD_PACK_SLS As Int32 = Val(rowICTITEM1.Item("ITEM_STD_PACK_SLS") & String.Empty)
            ' Allow Half Pack
            Dim ITEM_ALLOW_HALF_PACK As Boolean = (Val(rowICTITEM1.Item("ITEM_ALLOW_HALF_PACK") & String.Empty) = 1)

            ' Order Quantity meets Min Qty and Order Multiple restictions
            If ORDR_QTY >= ITEM_SO_QTY_MIN Then
                If ORDR_QTY Mod ITEM_SO_QTY_MULT = 0 Then
                    Continue For
                Else
                    ' If Half pack then (as of 02/21/2017 per LM) allow multiples of half pack
                    If ITEM_ALLOW_HALF_PACK Then
                        If ITEM_SO_QTY_MULT Mod 2 = 0 _
                            AndAlso ITEM_SO_QTY_MULT > 0 _
                            AndAlso ORDR_QTY Mod (ITEM_SO_QTY_MULT / 2) = 0 Then
                            Continue For
                        End If
                    End If

                    ' If Half pack then only half pack, not a case and a half pack
                    'If ITEM_ALLOW_HALF_PACK AndAlso ORDR_QTY = ITEM_SO_QTY_MULT / 2 _
                    '        AndAlso ITEM_SO_QTY_MULT Mod 2 = 0 _
                    '        AndAlso ITEM_SO_QTY_MULT > 0 Then
                    '    Continue For
                    'End If
                    'If ITEM_ALLOW_HALF_PACK AndAlso ORDR_QTY = ITEM_STD_PACK_SLS / 2 _
                    '        AndAlso ITEM_STD_PACK_SLS Mod 2 = 0 _
                    '        AndAlso ITEM_STD_PACK_SLS > 0 Then
                    '    Continue For
                    'End If

                End If

            End If


            errorMsg = "Order (" & ORDR_NO & "/" & IIf(multiStore, CUST_STORE_NO, ORDR_LNO) & ") has an Invalid Order Qty for Item: " & ITEM_CODE _
                & ", Min Qty = " & ITEM_SO_QTY_MIN & ", Order Multiple = " & ITEM_SO_QTY_MULT

            If Not errorList.Contains(errorMsg) Then
                errorList.Add(errorMsg)
            End If

            If Not salesOrderList.Contains(ORDR_NO & "/" & ORDR_LNO) Then
                salesOrderList.Add(ORDR_NO & "/" & ORDR_LNO)
            End If
        Next

        If errorList.Count > 0 Then
            ValidateOrderQtys = String.Join(vbCrLf, errorList.ToArray)
        End If

    End Function

    Public Shared Function BuildValidationTable(ByVal tblSOTORDR2 As DataTable, Optional tblSOTORDRS As DataTable = Nothing) As DataTable

        Dim vTable As New DataTable

        For Each COLUMN_NAME As String In New String() _
            {"CUST_CODE", "CUST_STORE_NO", "ORDR_NO", "ORDR_LNO", "ITEM_CODE", "ORDR_QTY_OPEN", "ORDR_QTY"}
            vTable.Columns.Add(COLUMN_NAME, tblSOTORDR2.Columns(COLUMN_NAME).DataType)
        Next

        For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "ITEM_CODE")
            Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE")
            Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO")
            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

            If tblSOTORDRS IsNot Nothing Then
                Dim QTY_COL As String = "QTY_" & Format(ORDR_LNO, "000")
                For Each rowSOTORDRS As DataRow In tblSOTORDRS.Select("", "CUST_STORE_NO")
                    Dim rowSOTORDRV As DataRow = vTable.NewRow
                    rowSOTORDRV.Item("CUST_CODE") = CUST_CODE
                    rowSOTORDRV.Item("CUST_STORE_NO") = rowSOTORDRS.Item("CUST_STORE_NO")
                    rowSOTORDRV.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDRV.Item("ORDR_LNO") = ORDR_LNO
                    rowSOTORDRV.Item("ITEM_CODE") = rowSOTORDR2.Item("ITEM_CODE")
                    rowSOTORDRV.Item("ORDR_QTY_OPEN") = Val(rowSOTORDRS.Item(QTY_COL) & "")
                    rowSOTORDRV.Item("ORDR_QTY") = Val(rowSOTORDRS.Item(QTY_COL) & "")
                    vTable.Rows.Add(rowSOTORDRV)
                Next
            Else
                Dim rowSOTORDRV As DataRow = vTable.NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"CUST_CODE", "CUST_STORE_NO", "ORDR_NO", "ORDR_LNO", "ITEM_CODE", "ORDR_QTY_OPEN", "ORDR_QTY"}
                    rowSOTORDRV.Item(COLUMN_NAME) = rowSOTORDR2.Item(COLUMN_NAME)
                Next
                vTable.Rows.Add(rowSOTORDRV)
            End If

        Next

        Return vTable

    End Function
    Public Shared Function CreateEDIReverse850(ByRef clsASCBASE1 As ASCBASE1, _
                                               ByVal ORDR_GROUP_NO As String) As String

        Dim EDI_DOC_SEQ_NO As String = String.Empty

        If clsASCBASE1.dst.Tables.Contains("EDT850T1") Then
            clsASCBASE1.dst.Tables("EDT850T1").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT850T2").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT850T3").Rows.Clear()
        Else
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT850T1", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT850T2", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT850T3", "*")
        End If

        ASCMAIN1.sql = "SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN " _
                & "( " _
                & " SELECT DISTINCT SOTORDR2.ITEM_CODE " _
                & " FROM SOTORDR1, SOTORDR2" _
                & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
                & " AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                & " AND SOTORDR2.ORDR_QTY > 0" _
                & " )"
        Dim tblItems As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTITEM1")

        Dim CUST_CODE As String = clsASCBASE1.dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", "")(0).Item("CUST_CODE")

        ' New way to get main customer from EDTSLSP1 to EDTTRPM1
        ASCMAIN1.sql = "SELECT * FROM EDTSLSP1 WHERE CUST_CODE = '" & CUST_CODE & "'"
        Dim tblEDTSLSP1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If tblEDTSLSP1.Rows.Count = 0 Then
            Return String.Empty
        End If

        Dim rowEDTSLSP1 As DataRow = tblEDTSLSP1.Rows(0)
        Dim EDI_QUAL_850 As String = rowEDTSLSP1.Item("EDI_QUAL_850") & String.Empty
        Dim EDI_ID_850 As String = rowEDTSLSP1.Item("EDI_ID_850") & String.Empty

        ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '850' AND EDI_TP_ID = '" & EDI_ID_850 & "' AND EDI_TP_QUAL = '" & EDI_QUAL_850 & "'"
        Dim tblEDTTRPM1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If tblEDTTRPM1.Rows.Count = 0 Then
            Return String.Empty
        End If

        EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDT850T1.EDI_DOC_SEQ_NO")

        Dim rowSOTORDR0 As DataRow = Nothing

        If clsASCBASE1.dst.Tables.Contains("SOTORDR0") AndAlso clsASCBASE1.dst.Tables("SOTORDR0").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'").Length > 0 Then
            rowSOTORDR0 = clsASCBASE1.dst.Tables("SOTORDR0").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")(0)
        Else
            rowSOTORDR0 = ASCDATA1.GetDataRow("Select * from SOTORDR0 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
        End If

        Dim EDI_ORD_QTY As Int32 = Val(clsASCBASE1.dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "") & String.Empty)
        Dim NUMBER_CHARS_STORE As Int32 = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
        Dim NUMBER_CHARS_DC As Int32 = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)

        ' Ceaate the EDI 850s
        Dim rowEDT850T1 As DataRow = clsASCBASE1.dst.Tables("EDT850T1").NewRow
        rowEDT850T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
        'rowEDT850T1.Item("EDI_JRNL_NO") = String.Empty
        rowEDT850T1.Item("GEN_DOC_NO") = "9999999999"
        'rowEDT850T1.Item("EDI_ISA_NO") = String.Empty

        rowEDT850T1.Item("EDI_TP_QUAL") = tblEDTTRPM1.Rows(0).Item("EDI_TP_QUAL") & String.Empty
        rowEDT850T1.Item("EDI_TP_ID") = tblEDTTRPM1.Rows(0).Item("EDI_TP_ID") & String.Empty
        rowEDT850T1.Item("EDI_OUR_QUAL") = tblEDTTRPM1.Rows(0).Item("EDI_OUR_QUAL") & String.Empty
        rowEDT850T1.Item("EDI_OUR_ID") = tblEDTTRPM1.Rows(0).Item("EDI_OUR_ID") & String.Empty

        'rowEDT850T1.Item("EDI_CUSTOMER") = String.Empty
        'rowEDT850T1.Item("EDI_STORE") = String.Empty
        rowEDT850T1.Item("EDI_DEPARTMENT") = rowSOTORDR0.Item("ORDR_DEPT") & String.Empty
        rowEDT850T1.Item("EDI_PO_NO") = rowSOTORDR0.Item("ORDR_CUST_PO") & String.Empty
        rowEDT850T1.Item("EDI_START_DATE") = rowSOTORDR0.Item("ORDR_SHIP_DATE")
        rowEDT850T1.Item("EDI_END_DATE") = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
        'rowEDT850T1.Item("EDI_SHIP_DATE") = String.Empty
        'rowEDT850T1.Item("EDI_SHIP_DC") = String.Empty
        'rowEDT850T1.Item("EDI_CENTER_CODE") = String.Empty
        'rowEDT850T1.Item("EDI_SUPPLIER_NO") = String.Empty
        'rowEDT850T1.Item("EDI_PROMOTION") = String.Empty
        'rowEDT850T1.Item("EDI_BATCH_NO") = String.Empty

        ' Requested on 6/29/2015 by Lauren
        If TACMAIN1.IPLBMacysCustomerCodes.Contains(CUST_CODE) Then
            rowEDT850T1.Item("EDI_MERCH_TYPE") = "CSMTCS/FRAG"
        End If

        'rowEDT850T1.Item("EDI_FOB") = String.Empty
        'rowEDT850T1.Item("EDI_TERMS") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_TYPE") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_BASIS") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_RATE") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_DSCDAYS") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_NETDAYS") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_DESC") = String.Empty
        'rowEDT850T1.Item("EDI_TERM_DOM") = String.Empty
        rowEDT850T1.Item("EDI_PO_PURP") = "00"
        rowEDT850T1.Item("EDI_PO_DATE") = rowSOTORDR0.Item("ORDR_DATE")
        rowEDT850T1.Item("EDI_PO_TYPE") = "SA"
        'rowEDT850T1.Item("EDI_SHIPPER") = String.Empty
        rowEDT850T1.Item("EDI_RECEIVED_DATE") = rowSOTORDR0.Item("ORDR_DATE_RECD")
        'rowEDT850T1.Item("EDI_SHIP_ADDR_TYPE") = String.Empty
        'rowEDT850T1.Item("EDI_CURRENCY") = String.Empty
        rowEDT850T1.Item("EDI_PROCESS_IND") = "1"
        rowEDT850T1.Item("EDI_ARRIVAL_DATE") = rowSOTORDR0.Item("ORDR_ARRIVAL_DATE")
        rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") = rowSOTORDR0.Item("ORDR_LAST_ARRIVAL_DATE")

        ' Done Below
        'rowEDT850T1.Item("EDI_NO_OF_LINES") = String.Empty

        'rowEDT850T1.Item("EDI_PRICE_BRACKET_ID") = String.Empty
        'rowEDT850T1.Item("STORE_GLOBAL_LOCATION_NUMBER") = String.Empty
        'rowEDT850T1.Item("DC_GLOBAL_LOCATION_NUMBER") = String.Empty
        'rowEDT850T1.Item("EDI_APPOINTMENT") = String.Empty
        'rowEDT850T1.Item("EDI_CARTON") = String.Empty
        'rowEDT850T1.Item("EDI_CURR_CODE") = String.Empty
        'rowEDT850T1.Item("EDI_DEPT_DESC") = String.Empty
        rowEDT850T1.Item("EDI_ORD_QTY") = EDI_ORD_QTY
        'rowEDT850T1.Item("EDI_WEIGHT") = String.Empty
        rowEDT850T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
        rowEDT850T1.Item("CUST_CODE") = rowSOTORDR0.Item("CUST_CODE")
        rowEDT850T1.Item("INIT_DATE") = DateTime.Now
        rowEDT850T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowEDT850T1.Item("LAST_DATE") = DateTime.Now
        rowEDT850T1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        'rowEDT850T1.Item("EDI_PO_RELEASE_NO") = String.Empty
        'rowEDT850T1.Item("EDI_PO_TOTAL_AMT") = String.Empty
        'rowEDT850T1.Item("EDI_CHAIN") = String.Empty
        'rowEDT850T1.Item("EDI_FACILITY") = String.Empty
        'rowEDT850T1.Item("EDI_CONTRACT_NO") = String.Empty
        clsASCBASE1.dst.Tables("EDT850T1").Rows.Add(rowEDT850T1)

        Dim EDI_DTL_SEQ As Int32 = 0
        For Each rowICTITEM1 As DataRow In tblItems.Select("", "ITEM_CODE")
            Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE")
            Dim rowEDT850T2 As DataRow = clsASCBASE1.dst.Tables("EDT850T2").NewRow
            rowEDT850T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
            EDI_DTL_SEQ += 1
            rowEDT850T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
            rowEDT850T2.Item("GEN_DOC_NO") = "9999999999"
            'rowEDT850T2.Item("EDI_BRAND") = String.Empty
            rowEDT850T2.Item("EDI_ITEM") = ITEM_CODE
            'rowEDT850T2.Item("EDI_DIMENSION") = String.Empty
            'rowEDT850T2.Item("EDI_SIZE_DESC") = String.Empty
            rowEDT850T2.Item("EDI_UPC") = rowICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
            'rowEDT850T2.Item("EDI_SKU") = String.Empty
            rowEDT850T2.Item("EDI_PO4_UOM") = rowICTITEM1.Item("ITEM_UOM") & String.Empty
            rowEDT850T2.Item("EDI_PRICE") = Val(clsASCBASE1.dst.Tables("SOTORDR2").Compute("MAX(ORDR_UNIT_PRICE)", "ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
            rowEDT850T2.Item("EDI_TOTAL_QTY") = Val(clsASCBASE1.dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
            'rowEDT850T2.Item("EDI_ITEM_NAME") = String.Empty
            'rowEDT850T2.Item("EDI_START_DATE") = String.Empty
            'rowEDT850T2.Item("EDI_END_DATE") = String.Empty
            'rowEDT850T2.Item("EDI_PO4_QTY") = String.Empty
            rowEDT850T2.Item("EDI_ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & String.Empty
            'rowEDT850T2.Item("EDI_PRICE_UOM") = String.Empty
            rowEDT850T2.Item("EDI_EAN") = rowICTITEM1.Item("ITEM_EAN_CODE") & String.Empty
            'rowEDT850T2.Item("EDI_PRICE_BRACKET_ID") = String.Empty
            'rowEDT850T2.Item("EDI_PO4_INNER") = String.Empty
            'rowEDT850T2.Item("EDI_GTIN") = String.Empty
            'rowEDT850T2.Item("EDI_COLOR_CODE") = String.Empty
            'rowEDT850T2.Item("EDI_COLOR_NAME") = String.Empty
            'rowEDT850T2.Item("EDI_DIVISION") = String.Empty
            'rowEDT850T2.Item("EDI_LBL_CODE") = String.Empty
            'rowEDT850T2.Item("EDI_STYLE") = String.Empty
            'rowEDT850T2.Item("EDI_STYLE_NAME") = String.Empty
            'rowEDT850T2.Item("EDI_RETAIL_PRICE") = String.Empty
            'rowEDT850T2.Item("EDI_TOTAL_AMT") = String.Empty
            'rowEDT850T2.Item("EDI_CARTON_GRP") = String.Empty
            'rowEDT850T2.Item("EDI_PO_LNO") = String.Empty

            ' Lord and taylor change 10/17/2018
            Dim EDI_GTIN As String = rowICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
            If EDI_GTIN.Length = 0 Then
                EDI_GTIN = rowICTITEM1.Item("ITEM_EAN_CODE") & String.Empty
            End If

            If EDI_GTIN.Length > 0 Then
                EDI_GTIN = EDI_GTIN.PadLeft(14, "0")
                rowEDT850T2.Item("EDI_GTIN") = EDI_GTIN
            End If

            clsASCBASE1.dst.Tables("EDT850T2").Rows.Add(rowEDT850T2)
        Next

        rowEDT850T1.Item("EDI_NO_OF_LINES") = clsASCBASE1.dst.Tables("EDT850T2").Rows.Count

        Dim EDI_SDQ_SEQ As Int32 = 0
        Dim rowEDT850T3 As DataRow = Nothing
        Dim storeMaxLen As Int32 = 0
        storeMaxLen = NUMBER_CHARS_STORE

        For Each rowEDT850T2 As DataRow In clsASCBASE1.dst.Tables("EDT850T2").Select("", "EDI_DTL_SEQ")
            EDI_DTL_SEQ = rowEDT850T2.Item("EDI_DTL_SEQ")
            EDI_SDQ_SEQ = 0

            Dim EDI_ITEM As String = rowEDT850T2.Item("EDI_ITEM")
            Dim fieldNum As Int16 = 0

            For Each rowSOTORDR2 As DataRow In clsASCBASE1.dst.Tables("SOTORDR2").Select("ITEM_CODE = '" & EDI_ITEM & "' AND ORDR_QTY > 0", "CUST_STORE_NO")
                Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = clsASCBASE1.dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                Dim EDI_STORE As String = String.Empty
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")

                EDI_STORE = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty

                If storeMaxLen > 0 And EDI_STORE.Length > storeMaxLen Then
                    EDI_STORE = StrReverse(StrReverse(EDI_STORE).Substring(0, storeMaxLen))
                End If

                If IsNumeric(EDI_STORE) AndAlso storeMaxLen > 0 Then
                    EDI_STORE = EDI_STORE.PadLeft(storeMaxLen, "0")
                End If

                fieldNum += 1

                Select Case fieldNum
                    Case 1
                        rowEDT850T3 = clsASCBASE1.dst.Tables("EDT850T3").NewRow
                        rowEDT850T3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                        rowEDT850T3.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                        EDI_SDQ_SEQ += 1
                        rowEDT850T3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                        rowEDT850T3.Item("GEN_DOC_NO") = "9999999999"
                        rowEDT850T3.Item("EDI_SDQ_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                        'rowEDT850T3.Item("EDI_SDQ_QUAL") = String.Empty
                        clsASCBASE1.dst.Tables("EDT850T3").Rows.Add(rowEDT850T3)
                End Select

                rowEDT850T3.Item("EDI_STORE_" & fieldNum.ToString("00")) = EDI_STORE
                rowEDT850T3.Item("EDI_QTY_" & fieldNum.ToString("00")) = rowSOTORDR2.Item("ORDR_QTY")

                If fieldNum = 10 Then
                    fieldNum = 0
                End If

                ' Update values in SOTORDR2
                rowSOTORDR2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                rowSOTORDR2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
            Next
        Next

        For Each tableName As String In New String() {"EDT850T1", "EDT850T2", "EDT850T3"}
            clsASCBASE1.Update_Record_TDA(tableName)
        Next

        ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE CUST_CODE = '" & CUST_CODE & "'"
        tblEDTTRPM1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        Dim ORDR_EDI_810 As String = IIf(tblEDTTRPM1.Select("EDI_DOC_NO = '810'").Length > 0, "1", "0")
        Dim ORDR_EDI_856 As String = IIf(tblEDTTRPM1.Select("EDI_DOC_NO = '856'").Length > 0, "1", "0")

        For Each rowSOTORDR1 As DataRow In clsASCBASE1.dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", "ORDR_NO")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            ASCMAIN1.sql = "UPDATE SOTORDR1 SET REVERSE_PO = '1', ORDR_SOURCE = 'E', ORDR_EDI_810 = '" & ORDR_EDI_810 & "'" _
                & ", ORDR_EDI_856 = '" & ORDR_EDI_856 & "', EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' WHERE ORDR_NO = '" & ORDR_NO & "'"
            ABSolution.ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            For Each rowSOTORDR2 As DataRow In clsASCBASE1.dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                ASCMAIN1.sql = "UPDATE SOTORDR2 SET EDI_DOC_SEQ_NO = '" & rowSOTORDR2.Item("EDI_DOC_SEQ_NO") & "'" _
                    & ", EDI_DTL_SEQ = '" & rowSOTORDR2.Item("EDI_DTL_SEQ") & "'" _
                    & " WHERE ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & rowSOTORDR2.Item("ORDR_LNO")
                ABSolution.ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Next
        Next

        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        clsASCBASE1.dst.Tables("EDT850T1").Rows.Clear()
        clsASCBASE1.dst.Tables("EDT850T2").Rows.Clear()
        clsASCBASE1.dst.Tables("EDT850T3").Rows.Clear()

        Return EDI_DOC_SEQ_NO

    End Function

    Public Shared Sub Create_PO_Rec_From_Return(frmASFBASE0 As ASFBASE0)

        Dim PO_ORDER_NO_rtn_KSP As String = "502672" ' "502647"
        Dim PO_ORDER_NO_rtn As String = "504549"

        'multitask the po
        Dim rowPOTORDR1 As DataRow = frmASFBASE0.Fill_Record("POTORDR1", PO_ORDER_NO_rtn)

        ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO_rtn & "'"
        frmASFBASE0.Fill_Records("POTORDR2", "", True, ASCMAIN1.sql)

        For Each rowSOTRTRN1 As DataRow In frmASFBASE0.dst.Tables("SOTRTRN1").Select("RTRN_AS_PO_REC = '1'")
            Dim RTRN_NO As String = rowSOTRTRN1.Item("RTRN_NO")

            frmASFBASE0.dst.Tables("APTINVH1").Rows.Clear()
            frmASFBASE0.dst.Tables("APTINVH2").Rows.Clear()
            frmASFBASE0.dst.Tables("APTINVH5").Rows.Clear()
            frmASFBASE0.dst.Tables("APTINVH5_VAR").Rows.Clear()

            frmASFBASE0.dst.Tables("ICTIREC1").Rows.Clear()
            frmASFBASE0.dst.Tables("ICTIREC2").Rows.Clear()
            Dim rowICTIREC1 As DataRow = Create_ICTIREC1(frmASFBASE0, rowPOTORDR1, rowSOTRTRN1)
            frmASFBASE0.dst.Tables("ICTIREC1").Rows.Add(rowICTIREC1)

            frmASFBASE0.dst.Tables("ICTPINV1").Rows.Clear()
            frmASFBASE0.dst.Tables("ICTPINV2").Rows.Clear()

            Dim rowICTPINV1 As DataRow = frmASFBASE0.dst.Tables("ICTPINV1").NewRow
            Dim PINV_NO As String = "0000000000"
            rowICTPINV1.Item("PINV_NO") = PINV_NO
            rowICTPINV1.Item("INV_NUM") = rowSOTRTRN1.Item("RTRN_NO")
            rowICTPINV1.Item("INV_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
            rowICTPINV1.Item("PINV_SOURCE_DOC") = ""
            rowICTPINV1.Item("RECEIPT_NO") = rowICTIREC1.Item("RECEIPT_NO")
            frmASFBASE0.dst.Tables("ICTPINV1").Rows.Add(rowICTPINV1)
            For Each rowSOTRTRN2 As DataRow In frmASFBASE0.dst.Tables("SOTRTRN2").Select("RTRN_NO = '" & RTRN_NO & "'")
                Dim ITEM_CODE As String = rowSOTRTRN2.Item("ITEM_CODE")
                Dim RTRN_QTY_2 As Integer = Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "") ' NEED TO INCLUDE RTRN_QTY_3 SINCE THIS WILL BE ADJUSTED OUT, SO MUST GO IN FIRST
                Dim rowICTITEM1 As DataRow = frmASFBASE0.LookUp("ICTITEM1", ITEM_CODE)
                Dim rowPOTORDR2 As DataRow = Nothing
                Dim rowPOTORDR2s() As DataRow = frmASFBASE0.dst.Tables("POTORDR2").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                If rowPOTORDR2s.Length > 0 Then
                    rowPOTORDR2 = rowPOTORDR2s(0)
                Else
                    rowPOTORDR2 = frmASFBASE0.dst.Tables("POTORDR2").NewRow
                    Dim PO_ORDER_LNO As Integer = Val(frmASFBASE0.dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", $"PO_ORDER_NO = '{PO_ORDER_NO_rtn}'")) + 1

                    With rowPOTORDR2
                        .Item("PO_ORDER_NO") = PO_ORDER_NO_rtn
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("ITEM_CODE") = ITEM_CODE
                        .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                        .Item("ITEM_UOM") = rowICTITEM1.Item("ITEM_UOM")
                        .Item("ITEM_PCT_ALLOW_OVER") = rowICTITEM1.Item("ITEM_PCT_ALLOW_OVER")
                        .Item("ITEM_PCT_ALLOW_UNDER") = rowICTITEM1.Item("ITEM_PCT_ALLOW_UNDER")
                        .Item("PO_COST") = rowSOTRTRN2.Item("ITEM_COST_STD")
                        .Item("PO_STATUS") = "O"
                        .Item("PO_DATE_REQUIRED") = rowPOTORDR1.Item("PO_DATE_REQUIRED")
                        .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                    End With
                    frmASFBASE0.dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
                End If

                rowPOTORDR2.Item("PO_COST") = rowSOTRTRN2.Item("ITEM_COST_STD")

                rowPOTORDR2.Item("PO_QTY_ORD") = Val(rowPOTORDR2.Item("PO_QTY_ORD") & "") + RTRN_QTY_2
                rowPOTORDR2.Item("PO_QTY_OPN") = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "") + RTRN_QTY_2

                'create ictirec2
                Dim rowICTIREC2 As DataRow = Create_ICTIREC2(frmASFBASE0, RTRN_QTY_2, rowICTIREC1, rowICTITEM1, rowPOTORDR1, rowPOTORDR2, rowSOTRTRN1)
                frmASFBASE0.dst.Tables("ICTIREC2").Rows.Add(rowICTIREC2)

                'create ictPINV2
                Dim rowICTPINV2 As DataRow = frmASFBASE0.dst.Tables("ICTPINV2").NewRow
                With rowICTPINV2
                    .Item("PINV_NO") = PINV_NO
                    .Item("PINV_LNO") = rowICTIREC2.Item("RECEIPT_LNO")
                    .Item("RECEIPT_NO") = rowICTIREC2.Item("RECEIPT_NO")
                    .Item("RECEIPT_LNO") = rowICTIREC2.Item("RECEIPT_LNO")
                    .Item("PINV_QTY") = rowICTIREC2.Item("QTY_REC")
                    .Item("PINV_COST") = rowICTIREC2.Item("PO_COST")
                    .Item("QTY_INV") = rowICTIREC2.Item("QTY_REC")
                End With
                frmASFBASE0.dst.Tables("ICTPINV2").Rows.Add(rowICTPINV2)
            Next

            TAC.POCMAIN1.ICTSTAT2_PO(-1, PO_ORDER_NO_rtn)
            frmASFBASE0.Update_Record_TDA("POTORDR2")
            TAC.POCMAIN1.ICTSTAT2_PO(1, PO_ORDER_NO_rtn)

            ICCMAIN1.Update_Receipt(frmASFBASE0, False)
            TAC.SOCMAIN1.Record_AP_Invoice(frmASFBASE0, rowICTIREC1, False, True, RTRN_NO)
        Next

    End Sub
    Public Shared Function Create_ICTIREC1(frmASFBASE0 As ASFBASE0, rowPOTORDR1 As DataRow, rowSOTRTRN1 As DataRow) As DataRow
        Dim rowICTIREC1 As DataRow = frmASFBASE0.dst.Tables("ICTIREC1").NewRow
        Dim DATETIME_STAMP As DateTime = DateTime.Now
        With rowICTIREC1
            .Item("RECEIPT_NO") = ASCMAIN1.Next_Control_No("ICTIREC1.RECEIPT_NO")
            .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE") & ""
            .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE") & ""
            .Item("RECEIPT_DATE") = DATETIME_STAMP.Date
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("REGISTER_IND") = "0"
            .Item("JOURNAL_IND") = "0"
            .Item("ACCRUAL_STATUS") = "0"
            .Item("PO_ORDER_NO") = rowPOTORDR1.Item("PO_ORDER_NO") & ""
        End With
        Return rowICTIREC1
    End Function

    Public Shared Function Create_ICTIREC2(frmASFBASE0 As ASFBASE0, QTY_REC As Int32, rowICTIREC1 As DataRow, rowICTITEM1 As DataRow _
                                           , rowPOTORDR1 As DataRow, rowPOTORDR2 As DataRow, rowSOTRTRN1 As DataRow) As DataRow

        Dim RECEIPT_LNO As Integer = Val(frmASFBASE0.dst.Tables("ICTIREC2").Compute("MAX(RECEIPT_LNO)", "") & $"RECEIPT_NO = '{rowICTIREC1.Item("RECEIPT_NO")}'") + 1
        Dim rowICTIREC2 As DataRow = frmASFBASE0.dst.Tables("ICTIREC2").NewRow
        Dim DATETIME_STAMP As DateTime = DateTime.Now
        With rowICTIREC2
            .Item("RECEIPT_NO") = rowICTIREC1.Item("RECEIPT_NO")
            .Item("RECEIPT_LNO") = RECEIPT_LNO
            .Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE")
            .Item("QTY_REC") = QTY_REC ' rowPOTORDR2.Item("PO_QTY_OPN")
            .Item("PO_ORDER_NO") = rowPOTORDR2.Item("PO_ORDER_NO")
            .Item("PO_ORDER_LNO") = rowPOTORDR2.Item("PO_ORDER_LNO")
            .Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD")
            .Item("PO_COST") = rowPOTORDR2.Item("PO_COST")
            .Item("ITEM_UOM") = rowPOTORDR2.Item("ITEM_UOM")

            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
            .Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
            .Item("REC_REF") = rowSOTRTRN1.Item("CUST_CODE")

            If rowPOTORDR2.Item("BM_ISSUE_SEL") & "" = "1" Or rowPOTORDR2.Item("BM_ISSUE_NO") & "" <> "" Then
                .Item("VEND_WHSE_CODE") = rowPOTORDR1.Item("VEND_WHSE_CODE")
                .Item("BM_ISSUE_SEL") = rowPOTORDR2.Item("BM_ISSUE_SEL")
                .Item("BM_ISSUE_NO") = rowPOTORDR2.Item("BM_ISSUE_NO")
            End If
        End With
        Return rowICTIREC2
    End Function

    Public Shared Sub Record_AP_Invoice(frmASFBASE0 As ASFBASE0, rowICTIREC1 As DataRow, reversal_update As Boolean, Optional as_po_rec As Boolean = False, Optional RTRN_NO As String = "")

        Dim DATETIME_STAMP As DateTime = DateTime.Now

        Dim VOUCHER_NO As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        Dim RECEIPT_NO As String = rowICTIREC1.Item("RECEIPT_NO")
        Dim VEND_CODE As String = rowICTIREC1.Item("VEND_CODE")
        Dim rowICTPINV1 As DataRow = frmASFBASE0.dst.Tables("ICTPINV1").Rows(0)
        Dim rowAPTVEND1 As DataRow = frmASFBASE0.LookUp("APTVEND1", VEND_CODE)

        If as_po_rec Then

        Else
            If reversal_update Then
                rowICTPINV1.Item("PINV_STATUS") = "O"
                rowICTPINV1.Item("VOUCHER_NO") = DBNull.Value
                rowICTPINV1.Item("RECEIPT_NO") = DBNull.Value
            Else
                rowICTPINV1.Item("PINV_STATUS") = "I"
                rowICTPINV1.Item("VOUCHER_NO") = VOUCHER_NO
                rowICTPINV1.Item("RECEIPT_NO") = RECEIPT_NO
            End If
        End If

        Dim INV_AMT As Decimal = Val(frmASFBASE0.dst.Tables("ICTPINV2").Compute("SUM(AMT_INV)", "") & "")

        If as_po_rec Then
            INV_AMT = 0 ' Returns as PO Receipts should be set up with a $0 invoice
        End If


        If reversal_update Then INV_AMT = -1 * INV_AMT

        Dim rowAPTINVH1 As DataRow = frmASFBASE0.dst.Tables("APTINVH1").NewRow
        With rowAPTINVH1
            .Item("VOUCHER_NO") = VOUCHER_NO
            .Item("VEND_CODE") = VEND_CODE
            .Item("INV_TYPE") = "I"
            .Item("INV_NUM") = rowICTPINV1.Item("INV_NUM")
            .Item("INV_DATE") = rowICTPINV1.Item("INV_DATE")

            .Item("INV_AMT") = INV_AMT
            .Item("INV_REF") = rowICTPINV1.Item("PINV_SOURCE_DOC")

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
                .Item("BANK_CODE") = frmASFBASE0.ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
            Else
                .Item("BANK_CODE") = rowAPTVEND1.Item("BANK_CODE")
            End If

            If rowAPTVEND1.Item("VEND_PYMT_METHOD") & "" = "" Then
                If .Item("BANK_CODE") & "" <> "" Then
                    Dim rowGLTBANK1 As DataRow = frmASFBASE0.LookUp("GLTBANK1", .Item("BANK_CODE"))
                    .Item("INV_PYMT_METHOD") = rowGLTBANK1.Item("BANK_PYMT_METHOD")
                End If
            Else
                .Item("INV_PYMT_METHOD") = rowAPTVEND1.Item("VEND_PYMT_METHOD")
            End If

            .Item("INV_PYMT_CYCLE") = rowAPTVEND1.Item("VEND_PYMT_CYCLE")

            If rowAPTVEND1.Item("POST_CODE") & "" <> "" Then
                .Item("POST_CODE") = rowAPTVEND1.Item("POST_CODE")
            Else
                .Item("POST_CODE") = frmASFBASE0.ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
            End If

            .Item("INV_STATUS") = "O"
            .Item("INV_PYMT_CYCLE") = DBNull.Value
            .Item("INV_DUE_DATE") = TAC.TACMAIN1.Calculate_INV_DUE_DATE(frmASFBASE0, rowAPTVEND1.Item("TERM_CODE") & "", Nothing, .Item("INV_DATE"))
            .Item("INV_BALANCE") = INV_AMT
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("CURR_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1

            .Item("SEG2_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            .Item("REGISTER_IND") = "0"
            .Item("INV_BL_DATE") = .Item("INV_DATE")
            .Item("INV_AMT_VEND") = .Item("INV_AMT")

            If rowAPTVEND1.Item("VEND_AUTO_APPROVE") & "" = "1" Then
                .Item("INV_APPR_STATUS") = "A"
                frmASFBASE0.Write_Event_Log("APTINVH1", VOUCHER_NO, "Auto Approved")
            Else
                .Item("INV_APPR_STATUS") = "P"
            End If

            .Item("VEND_BUYER_CODE") = rowAPTVEND1.Item("VEND_BUYER_CODE")
        End With

        frmASFBASE0.dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)

        ASCMAIN1.Record_Event("APTINVH1", VOUCHER_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "AUTO", "Auto Rec/Inv", RECEIPT_NO)

        ' ADO.NET BEHAVIOR USING CUSTOM SQL STATEMENTS:
        ' NOTE THAT THE FIELDS VAR_QTY AND VAR_AMT ARE NOT ACCOUNTED FOR IN THE SQL
        ' THEY SHOULD BE RIGHT BEFORE '0' VB
        ' YET THE FILL MATCHES BASED ON NAME SO THAT CB AND COST_CATGY_CODE GET THE VALUES COMING IN FROM THOSE COLUMN NAMES
        ' MEANING NAME MAPPING TRUMPS POSITIONAL MAPPING

        '            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) INV_QTY" & vbCrLf _
        '            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _

        Dim sql As String = "Select '" & VOUCHER_NO & "' VOUCHER_NO " & vbCrLf _
            & ", ICTIREC2.RECEIPT_LNO VOUCHER_DLNO " & vbCrLf _
            & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,0) INV_QTY" & vbCrLf _
            & ", ICTIREC2.PO_COST INV_COST" & vbCrLf _
            & ", '0' CB" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE, '0' CLOSE_LINE" & vbCrLf _
            & ", DECODE(ICTIREC2.ACCRUAL_STATUS,'1',0,NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0)) QTY_REC_NOT_INV" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOST1.ACCT_CODE_PPV" & vbCrLf _
            & ", ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.PO_COST " & vbCrLf _
            & ", ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM" & vbCrLf _
            & " from ICTIREC2,ICTITEM1,ICTCOST1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & " and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE"
        frmASFBASE0.Fill_Records("APTINVH5", , False, sql)

        frmASFBASE0.dst.Tables("APTINVH5").AcceptChanges()
        For Each rowAPTINVH5 As DataRow In frmASFBASE0.dst.Tables("APTINVH5").Select("")
            rowAPTINVH5.SetAdded()
            Dim RECEIPT_LNO As Int64 = Val(rowAPTINVH5.Item("RECEIPT_LNO") & "")
            Dim rowICTIREC2 As DataRow = frmASFBASE0.dst.Tables("ICTIREC2").Rows.Find(New Object() {RECEIPT_NO, RECEIPT_LNO})
            Dim rowICTPINV2 As DataRow = frmASFBASE0.dst.Tables("ICTPINV2").Select("RECEIPT_LNO = " & CStr(RECEIPT_LNO))(0)
            Dim QTY_INV As Int64 = Val(rowICTPINV2.Item("QTY_INV") & "")

            If reversal_update Then
                QTY_INV = -1 * QTY_INV
            End If

            Dim PINV_COST As Decimal = rowICTPINV2.Item("PINV_COST")
            Dim VAR_AMT As Decimal = 0
            rowAPTINVH5.Item("INV_QTY") = QTY_INV
            rowAPTINVH5.Item("INV_COST") = PINV_COST
            ' NEXT 2 LINES ADDED BELOW TO SUPPORT THEM - THEY WERE NULL IN ORACLE
            ' VARIANCE WAS BEING CALCULATED And TRANSACTED TO GL PROPERLY 
            ' BECAUSE EXTENDED FIELDS QTY_VAR AND AMT_VAR WERE BEING CALCULATED AND USED
            rowAPTINVH5.Item("VAR_QTY") = 0 ' Val(rowAPTINVH5.Item("INV_QTY") & "") - Val(rowICTIREC2.Item("QTY_REC") & "")
            rowAPTINVH5.Item("VAR_AMT") = (Val(rowAPTINVH5.Item("INV_COST") & "") - Val(rowICTIREC2.Item("PO_COST") & "")) * Val(rowAPTINVH5.Item("INV_QTY") & "")
        Next

        Create_APTINVH5_VAR(frmASFBASE0)
        Create_APTINVH2_P(frmASFBASE0, rowAPTINVH1, rowICTIREC1, as_po_rec, RTRN_NO)

        frmASFBASE0.Update_Record_TDA("APTINVH1")
        frmASFBASE0.Update_Record_TDA("APTINVH2")
        frmASFBASE0.Update_Record_TDA("APTINVH5")
        If Not as_po_rec Then
            frmASFBASE0.Update_Record_TDA("ICTPINV1")
            frmASFBASE0.Update_Record_TDA("ICTPINV2")
        End If

        Dependent_Updates(frmASFBASE0, rowAPTINVH1, False)

    End Sub

    Public Shared Sub Create_APTINVH5_VAR(frmASFBASE0 As ASFBASE0)

        frmASFBASE0.dst.Tables("APTINVH5_VAR").Rows.Clear()

        For Each rowAPTINVH5 As DataRow In frmASFBASE0.dst.Tables("APTINVH5") _
        .Select("", "", DataViewRowState.CurrentRows)
            Dim COST_CATGY_CODE As String = rowAPTINVH5.Item("COST_CATGY_CODE")
            Dim rowICTCOST1 As DataRow = frmASFBASE0.LookUp("ICTCOST1", COST_CATGY_CODE)

            Dim COLLECTION_CODE As String = rowAPTINVH5.Item("COLLECTION_CODE")
            Dim rowICTCOLL1 As DataRow = frmASFBASE0.LookUp("ICTCOLL1", COLLECTION_CODE) '  dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE) ' LookUp("ICTCOLL1", COLLECTION_CODE)

            Dim SEG2_CODE As String = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            Dim SEG3_CODE As String = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            Dim SEG4_CODE As String = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            If frmASFBASE0.ROWs("ICTPARM1").Item("IC_PARM_EXP_SEG4") & "" = "1" Then
                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                    SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                Else
                    SEG4_CODE = COLLECTION_CODE
                End If
            End If

            Dim rowAPTINVH5_VAR As DataRow = frmASFBASE0.dst.Tables("APTINVH5_VAR").Rows.Find(New String() {COST_CATGY_CODE, COLLECTION_CODE})
            If rowAPTINVH5_VAR Is Nothing Then
                rowAPTINVH5_VAR = frmASFBASE0.dst.Tables("APTINVH5_VAR").NewRow
                rowAPTINVH5_VAR.Item("COST_CATGY_CODE") = COST_CATGY_CODE
                rowAPTINVH5_VAR.Item("COLLECTION_CODE") = COLLECTION_CODE
                rowAPTINVH5_VAR.Item("ACCT_CODE_PPV") = rowAPTINVH5.Item("ACCT_CODE_PPV")
                Dim rowICTIREC1 As DataRow = frmASFBASE0.dst.Tables("ICTIREC1").Rows.Find(rowAPTINVH5.Item("RECEIPT_NO"))
                rowAPTINVH5_VAR.Item("SEG2_CODE") = SEG2_CODE ' rowICTIREC1.Item("SEG2_CODE")
                rowAPTINVH5_VAR.Item("SEG3_CODE") = SEG3_CODE
                rowAPTINVH5_VAR.Item("SEG4_CODE") = SEG4_CODE
                frmASFBASE0.dst.Tables("APTINVH5_VAR").Rows.Add(rowAPTINVH5_VAR)
            End If
            rowAPTINVH5_VAR.Item("AMT_REC") = Val(rowAPTINVH5_VAR.Item("AMT_REC") & "") + Val(rowAPTINVH5.Item("AMT_REC") & "")
            rowAPTINVH5_VAR.Item("AMT_INV") = Val(rowAPTINVH5_VAR.Item("AMT_INV") & "") + Val(rowAPTINVH5.Item("AMT_INV") & "")
            rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV") = Val(rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV") & "") + Val(rowAPTINVH5.Item("AMT_REC_NOT_INV") & "") ' Val(rowAPTINVH5.Item("QTY_REC_NOT_INV") & "") * Val(rowAPTINVH5.Item("PO_COST") & "")
            rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV_OFFSET") = Val(rowAPTINVH5_VAR.Item("AMT_REC_NOT_INV_OFFSET") & "") + Val(rowAPTINVH5.Item("AMT_REC_NOT_INV_OFFSET") & "")
            rowAPTINVH5_VAR.Item("AMT_VAR") = Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            If rowAPTINVH5.Item("CB") & "" = "1" Then
                rowAPTINVH5_VAR.Item("AMT_VAR_CB") = Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            End If
        Next

    End Sub

    Public Shared Sub Create_APTINVH2_P(frmASFBASE0 As ASFBASE0, rowAPTINVH1 As DataRow,
                                        rowICTIREC1 As DataRow, as_po_rec As Boolean, Optional RTRN_NO As String = "")

        Dim ACCT_CODE_RSV2 As String = "222250"
        Dim ACCT_CODE_X As String = "152600"

        frmASFBASE0.Delete_Rows("APTINVH2", "INV_LTYP = 'P'")

        Dim VOUCHER_NO As String = rowAPTINVH1.Item("VOUCHER_NO")
        Dim RECEIPT_NO As String = rowICTIREC1.Item("RECEIPT_NO")
        Dim VOUCHER_LNO_ctr As Int32 = Val(frmASFBASE0.dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & VOUCHER_NO & "'") & "")

        Dim INV_LINE_AMT As Decimal = Val(frmASFBASE0.dst.Tables("APTINVH5_VAR").Compute("SUM(AMT_REC_NOT_INV_OFFSET)", "") & "")
        Dim REV As Decimal = Val(frmASFBASE0.dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES)", "RTRN_NO = '" & RTRN_NO & "' and RTRN_AS_PO_REC = '1'") & "")
        Dim COST As Decimal = 0 'Val(frmASFBASE0.dst.Tables("ICTIREC2").Compute("SUM(PO_COST * QTY_REC)", "RECEIPT_NO = '" & RECEIPT_NO & "'") & "")

        For Each ROW As DataRow In frmASFBASE0.dst.Tables("ICTIREC2").Select("RECEIPT_NO = '" & RECEIPT_NO & "'")
            COST += (Val(ROW.Item("PO_COST") & "") * Val(ROW.Item("QTY_REC") & ""))
        Next

        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = frmASFBASE0.dst.Tables("APTINVH2").NewRow
            With rowAPTINVH2
                .Item("VOUCHER_NO") = VOUCHER_NO
                VOUCHER_LNO_ctr += 1
                .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                .Item("ACCT_CODE") = frmASFBASE0.ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_PURCH")
                .Item("SEG2_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("INV_LINE_AMT") = INV_LINE_AMT
                .Item("INV_LTYP") = "P"
            End With
            frmASFBASE0.dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        End If


        If as_po_rec Then
            rowAPTINVH2 = frmASFBASE0.dst.Tables("APTINVH2").NewRow
            With rowAPTINVH2
                .Item("VOUCHER_NO") = VOUCHER_NO
                VOUCHER_LNO_ctr += 1
                .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                .Item("ACCT_CODE") = ACCT_CODE_RSV2
                .Item("SEG2_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("INV_LINE_AMT") = REV - COST
            End With
            frmASFBASE0.dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)

            rowAPTINVH2 = frmASFBASE0.dst.Tables("APTINVH2").NewRow
            With rowAPTINVH2
                .Item("VOUCHER_NO") = VOUCHER_NO
                VOUCHER_LNO_ctr += 1
                .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                .Item("ACCT_CODE") = ACCT_CODE_X
                .Item("SEG2_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                .Item("INV_LINE_AMT") = REV * -1
            End With
            frmASFBASE0.dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        End If

        For Each rowAPTINVH5_VAR As DataRow In frmASFBASE0.dst.Tables("APTINVH5_VAR").Select("ISNULL(AMT_VAR,0) - ISNULL(AMT_VAR_CB,0) <> 0")
            ' INV_LINE_AMT = -1 * (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            INV_LINE_AMT = (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            If INV_LINE_AMT <> 0 Then
                rowAPTINVH2 = frmASFBASE0.dst.Tables("APTINVH2").NewRow
                With rowAPTINVH2
                    .Item("VOUCHER_NO") = VOUCHER_NO
                    VOUCHER_LNO_ctr += 1
                    .Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                    .Item("ACCT_CODE") = rowAPTINVH5_VAR.Item("ACCT_CODE_PPV")
                    .Item("SEG2_CODE") = rowAPTINVH5_VAR.Item("SEG2_CODE")
                    .Item("SEG3_CODE") = rowAPTINVH5_VAR.Item("SEG3_CODE")
                    .Item("SEG4_CODE") = rowAPTINVH5_VAR.Item("SEG4_CODE")
                    .Item("INV_LINE_AMT") = INV_LINE_AMT
                    .Item("INV_LTYP") = "P"
                End With
                frmASFBASE0.dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If
        Next
    End Sub

    Public Shared Sub Dependent_Updates(frmASFBASE0 As ASFBASE0, ByVal rowAPTINVH1 As DataRow, ByVal reverse As Boolean)

        Dim VOUCHER_NO As String = rowAPTINVH1.Item("VOUCHER_NO")
        Dim VEND_CODE As String = rowAPTINVH1.Item("VEND_CODE")
        Dim rowAPTVEND5 As DataRow = frmASFBASE0.Fill_Record("APTVEND5", VEND_CODE, True)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is Select * from APTINVH5 where VOUCHER_NO = '" & VOUCHER_NO & "';" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTIREC2 Set " & vbCrLf _
            & IIf(reverse,
                  "QTY_INV = NVL(QTY_INV,0) - NVL(R1.INV_QTY,0)",
                  "QTY_INV = NVL(QTY_INV,0) + NVL(R1.INV_QTY,0)") & vbCrLf _
            & IIf(reverse,
                  ",AMT_INV = NVL(AMT_INV,0) - NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)",
                  ",AMT_INV = NVL(AMT_INV,0) + NVL(R1.INV_QTY,0) * NVL(R1.INV_COST,0)") & vbCrLf _
            & IIf(reverse,
                  ", ACCRUAL_STATUS = '0'",
                  ", ACCRUAL_STATUS = NVL(R1.CLOSE_LINE,'0')") & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "   Update ICTIREC2 Set ACCRUAL_STATUS = '1'" & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO and NVL(QTY_REC,0) = NVL(QTY_INV,0);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select Distinct RECEIPT_NO from APTINVH5 where VOUCHER_NO = '" & VOUCHER_NO & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select()

            Dim RECEIPT_NO As String = row.Item("RECEIPT_NO")

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select RECEIPT_NO" & vbCrLf _
                & " , Min (ACCRUAL_STATUS) ACCRUAL_STATUS" & vbCrLf _
                & " , Sum (QTY_INV) QTY_INV, Sum (AMT_INV) AMT_INV" & vbCrLf _
                & "  from ICTIREC2 where RECEIPT_NO = '" & RECEIPT_NO & "' group by RECEIPT_NO;" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTIREC1 Set " & vbCrLf _
                & "    ACCRUAL_STATUS = R1.ACCRUAL_STATUS" & vbCrLf _
                & "    , QTY_INV = R1.QTY_INV" & vbCrLf _
                & "    , AMT_INV = R1.AMT_INV" & vbCrLf _
                & "    where RECEIPT_NO = R1.RECEIPT_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSP("ICPIRECV", "V", New Object() {RECEIPT_NO}, New String() {"RECEIPT_NO_in"})
        Next

        Dim INV_AMT As Decimal = Val(rowAPTINVH1("INV_AMT") & "")
        If reverse Then
            Dim row As DataRow = frmASFBASE0.LookUp("APTINVH1", VOUCHER_NO)
            INV_AMT = -1 * Val(row("INV_AMT") & "")
        End If
        ' rowAPTVEND5 = Fill_Record("APTVEND5", VEND_CODE, True)

        With rowAPTVEND5
            .Item("VEND_PURCHASES_MTD") = Val(.Item("VEND_PURCHASES_MTD") & "") + INV_AMT
            .Item("VEND_PURCHASES_YTD") = Val(.Item("VEND_PURCHASES_YTD") & "") + INV_AMT
            .Item("VEND_NUM_INV_MTD") = Val(.Item("VEND_NUM_INV_MTD") & "") + IIf(reverse, -1, 1)
            .Item("VEND_NUM_INV_YTD") = Val(.Item("VEND_NUM_INV_YTD") & "") + IIf(reverse, -1, 1)
            If Not reverse Then
                .Item("VEND_LAST_INV_DATE") = rowAPTINVH1("INV_DATE")
                .Item("VEND_LAST_INV_AMT") = rowAPTINVH1("INV_AMT")
                .Item("VEND_LAST_INV_NUM") = rowAPTINVH1("INV_NUM")
                If .Item("VEND_1ST_PURCH_DATE") & "" = "" Then
                    .Item("VEND_1ST_PURCH_DATE") = rowAPTINVH1("INV_DATE")
                End If
            End If

        End With
    End Sub

#Region "Evaluate 3PL Sales Orders"

    Public Enum evaluation_types
        sales_order
        pick_ticket

    End Enum

    Public Shared Function Evaluate_Order(ByVal invalue As String, ByVal ordrLNO As String, evaluation_type As evaluation_types) As List(Of String)
        Select Case evaluation_type
            Case evaluation_types.sales_order
                Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTORDR1 WHERE ORDR_NO = :PARM1", "V", {invalue})
                Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR2 WHERE ORDR_NO = :PARM1 and ORDR_LNO = :PARM2", "SOTORDR2", "VV", {invalue, ordrLNO})
                Return Evaluate_Order(rowSOTORDR1, tblSOTORDR2, evaluation_types.sales_order)


            Case evaluation_types.pick_ticket
                Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTPICK1 WHERE ORDR_NO = :PARM1", "V", {invalue})
                Dim tblSOTPICK2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTPICK2 WHERE ORDR_NO = :PARM1", "SOTPICK2", "V", {invalue})
                Return Evaluate_Order(rowSOTPICK1, tblSOTPICK2, evaluation_types.pick_ticket)

            Case Else
                Return New List(Of String)

        End Select
    End Function

    Public Shared Function Evaluate_Order(ByVal rowData As DataRow, ByVal tableData As DataTable, evaluation_type As evaluation_types) As List(Of String)

        'FOR SO ENTRY: TAC.SOCMAIN1.Evaluate_Order(rowSOTORDR1, tblSOTORDR2, SOCMAIN1.evaluation_types.sales_order)
        'FOR SO RELEASE: TAC.SOCMAIN1.Evaluate_Order(rowSOTORDR1, tblSOTPICK2, SOCMAIN1.evaluation_types.pick_ticket)

        Dim results As New List(Of String)

        'IF ROW IS NOTHING RETURN THE EMPTY LIST
        If rowData Is Nothing Then
            Return results
        End If

        'IF TABLE IS NOTHING OR EMPTY RETURN EMPTY LIST
        If tableData Is Nothing Or tableData.Rows.Count = 0 Then
            Return results
        End If

        'WE ARE INTERESTED IN THOSE 3PL WAREHOUSES (CLA, CLARTN, ETC.)
        Dim ORDR_NO As String = rowData.Item("ORDR_NO") & ""
        Dim WHSE_CODE As String
        If evaluation_type = evaluation_types.pick_ticket Then
            WHSE_CODE = ASCDATA1.GetDataValue($"SELECT WHSE_CODE FROM SOTORDR1 WHERE ORDR_NO = '{ORDR_NO}'")
        Else
            WHSE_CODE = rowData.Item("WHSE_CODE") & ""
        End If
        Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow("Select * from ICTWHSE1 WHERE WHSE_CODE = : PARM1", "V", {WHSE_CODE})

        'IF THE ROW IS NOTHING RETURN THE EMPTY LIST
        If rowICTWHSE1 Is Nothing Then
            Return results
        End If

        Dim MAX_DETAIL_LINE_QTY As Int32 = Val(rowICTWHSE1.Item("MAX_DETAIL_LINE_QTY") & "")
        Dim MAX_SALES_ORDER_QTY As Int32 = Val(rowICTWHSE1.Item("MAX_SALES_ORDER_QTY") & "")

        'IF BOTH ARE 0 OR LESS RETURN THE EMPTY LIST
        If MAX_DETAIL_LINE_QTY <= 0 And MAX_SALES_ORDER_QTY <= 0 Then
            Return results
        End If

        If MAX_DETAIL_LINE_QTY > 0 Then
            Dim FIELD_NAME As String
            Select Case evaluation_type
                Case evaluation_types.sales_order
                    FIELD_NAME = "ORDR_QTY"

                Case evaluation_types.pick_ticket
                    FIELD_NAME = "PICK_QTY"

                Case Else
                    Return results
            End Select

            For Each ROW As DataRow In tableData.Select()
                'IF ORDR QTY/PICK QTY IS MORE THAN THE  MAX (99,999 OR WHATEVER THEY SPECIFY) ADD A LINE TO EMSG
                If Val(ROW.Item(FIELD_NAME) & "") > MAX_DETAIL_LINE_QTY Then
                    results.Add($"Order Line Number {ROW.Item("ORDR_LNO")} exceeds maximum detail line quantity.")
                End If

            Next

        End If

        If MAX_SALES_ORDER_QTY > 0 Then
            Select Case evaluation_type
                Case evaluation_types.sales_order
                    If Val(tableData.Compute("Sum(ORDR_QTY_OPEN)", "ORDR_STATUS = 'O' OR ORDR_STATUS = 'P'") & "") > MAX_SALES_ORDER_QTY Then
                        results.Add($"Order total pieces exceeds maximum order quantity.")
                    End If

                Case evaluation_types.pick_ticket
                    If Val(tableData.Compute("Sum(PICK_QTY)", "") & "") > MAX_SALES_ORDER_QTY Then
                        results.Add($"Order total pieces exceeds maximum order quantity.")
                    End If
            End Select
        End If
        Return results
    End Function
#End Region

End Class
