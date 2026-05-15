Public Class ARCMAIN1

    Public Shared AGE_DAYS(4) As String
    Public Shared DUE_DAYS(4) As String
    Public Shared AGE_DATE(4) As String
    Public Shared DUE_DATE(4) As String
    Public Shared AGE_DATE_ORA(4) As String
    Public Shared DUE_DATE_ORA(4) As String
    Public Shared DAYS_AND_BUCKETS As String
    Public Shared AGED_TOTALS As String

    Public Shared Sub Record_Customer_Event(ByVal CUST_CODE As String, ByVal EVENT_DESC As String, ByVal EVENT_TYPE As String)
        TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, EVENT_TYPE, EVENT_DESC)
    End Sub

    Public Shared Sub Pay_Open_AR_Item( _
    ByVal rowARTOPEN1 As DataRow, _
    ByVal rowARTPYMT2 As DataRow, _
    ByVal CURR_CODE As String, _
    ByVal PYMT_BATCH_DATE As Date, _
    ByVal INV_PMT As Double, _
    ByVal INV_DISC_TAKEN As Double, _
    ByVal INV_WRITE_OFF As Double, _
    ByRef PYMT_BATCH_ILNO As Integer, _
    ByVal F As ASFBASE1)

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
            .Item("LAST_DATE") = F.DATETIME_STAMP
            If CURR_CODE = F.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                .Item("INV_PMT_CURR") = .Item("INV_PMT")
                .Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
                .Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            Else
                MsgBox("ABS Stop Statement - Currency not set up") ' Stop
            End If

            Dim rowARTPYMT3 As DataRow = F.dst.Tables("ARTPYMT3").NewRow

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
            If CURR_CODE = F.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT3.Item("INV_BALANCE_CURR") = rowARTPYMT3.Item("INV_BALANCE")
                rowARTPYMT3.Item("INV_PMT_CURR") = rowARTPYMT3.Item("INV_PMT")
                rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") = rowARTPYMT3.Item("INV_DISC_TAKEN")
                rowARTPYMT3.Item("INV_WRITE_OFF_CURR") = rowARTPYMT3.Item("INV_WRITE_OFF")
                rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") = rowARTPYMT3.Item("INV_BALANCE_NEW")
            Else
                MsgBox("ABS Stop Statement - Currency not set up") ' Stop
            End If
            F.dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
        End With
    End Sub
    Public Shared Function Last_3612_SAvg(ByVal CUST_CREDIT_GROUP_CUST As String) As String

        Dim LAST_3612 As String = ""
        Dim YP3 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        Dim YP6 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)
        Dim YP12 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)

        ASCMAIN1.sql = "SELECT MIN (OPS_YYYYPP) YP" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP3 & "' THEN INV_CLSD_DYS ELSE 0 END) D03" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP3 & "' THEN INV_CLSD_CNT ELSE 0 END) T03" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP3 & "' THEN INV_CLSD_CNT ELSE 0 END) C03" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP6 & "' THEN INV_CLSD_DYS ELSE 0 END) D06" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP6 & "' THEN INV_CLSD_CNT ELSE 0 END) T06" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP6 & "' AND OPS_YYYYPP <'" & YP3 & "' THEN INV_CLSD_CNT ELSE 0 END) C06" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP12 & "' THEN INV_CLSD_DYS ELSE 0 END) D12" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP12 & "' THEN INV_CLSD_CNT ELSE 0 END) T12" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP12 & "' AND OPS_YYYYPP <'" & YP6 & "' THEN INV_CLSD_CNT ELSE 0 END) C12" _
        & " from ARTSTMT1 where OPS_YYYYPP >= '" & YP12 & "'" _
        & "  and ARTSTMT1.CUST_CODE = '" & CUST_CREDIT_GROUP_CUST & "'"
        '& "  and ARTSTMT1.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")"

        Dim row3612 As DataRow = ASCDATA1.GetDataRow
        Dim YP As String = row3612.Item("YP") & ""
        Dim V As Decimal = 0
        Dim DXX As Decimal = 0
        Dim TXX As Decimal = 0
        Dim CXX As Integer = 0
        If YP <= YP3 Then
            DXX = Val(row3612.Item("D03") & "")
            TXX = Val(row3612.Item("T03") & "")
            CXX = Val(row3612.Item("C03") & "")
            If TXX = 0 Or CXX = 0 Then
                V = 0
            Else
                V = DXX / TXX
            End If
            LAST_3612 &= "/" & Format(V, "##0")
            If YP <= YP6 Then
                DXX = Val(row3612.Item("D06") & "")
                TXX = Val(row3612.Item("T06") & "")
                CXX = Val(row3612.Item("C06") & "")
                If TXX = 0 Or CXX = 0 Then
                    V = 0
                Else
                    V = DXX / TXX
                End If
                LAST_3612 &= "/" & Format(V, "##0")
                If YP <= YP12 Then
                    DXX = Val(row3612.Item("D12") & "")
                    TXX = Val(row3612.Item("T12") & "")
                    CXX = Val(row3612.Item("C12") & "")
                    If TXX = 0 Or CXX = 0 Then
                        V = 0
                    Else
                        V = DXX / TXX
                    End If
                    LAST_3612 &= "/" & Format(V, "##0")
                End If
            End If
        End If
        LAST_3612 = Mid(LAST_3612, 2)

        Return LAST_3612
    End Function
    Public Shared Sub Get_Aging_Data( _
         ByVal rowARTPARM1 As DataRow, _
         ByVal BASE_DATE As Date, _
         Optional ByVal use_parms_for_days As Boolean = True, _
         Optional ByVal consider_future As Boolean = False)

        If use_parms_for_days Then
            For i As Integer = 2 To 4
                AGE_DAYS(i) = Val(rowARTPARM1.Item("AR_PARM_AGE_CATG_" & CStr(i)) & "")
                DUE_DAYS(i) = Val(rowARTPARM1.Item("AR_PARM_DUE_CATG_" & CStr(i)) & "")
            Next
        End If
        For i As Integer = 1 To 4
            Dim PRD_END_DATE As Date = BASE_DATE.AddDays(-1 * AGE_DAYS(i))
            AGE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
            AGE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        Next
        For i As Integer = 1 To 4
            Dim PRD_END_DATE As Date = BASE_DATE.AddDays(-1 * DUE_DAYS(i))
            DUE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
            DUE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        Next

        'Dim sqlINV As String = "ARTOPEN1.INV_DATE"
        'Dim sqlDUE As String = "ARTOPEN1.INV_DUE_DATE"

        DAYS_AND_BUCKETS = "" _
        & ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "')) - ARTOPEN1.INV_DATE DAYS" & vbCrLf _
        & ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(2) & "                                                  THEN '1' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        & "  '4' END END END END) AGE_BUCKET" & vbCrLf _
        & ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DATE) AGE" & vbCrLf _
        & ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & "                                                     THEN '1' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        & "  '4' END END END END) DUE_BUCKET" & vbCrLf _
        & ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DUE_DATE) DUE"

        Dim sqle As String = " THEN ARTOPEN1.INV_BALANCE ELSE 0 END "
        '        & IIf(consider_future, ", 0 DUE_0" & vbCrLf, ", 0 DUE_0") _

        AGED_TOTALS = "" _
        & IIf(consider_future, _
          ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN1.AGE_DATE_ORA(1) & "                                                         " & sqle & ") AGE_0" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN1.AGE_DATE_ORA(2) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN1.AGE_DATE_ORA(1) & sqle & ") AGE_1" & vbCrLf, _
          ", 0 AGE_0" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN1.AGE_DATE_ORA(2) & "                                                         " & sqle & ") AGE_1" & vbCrLf) _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN1.AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN1.AGE_DATE_ORA(2) & sqle & ") AGE_2" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN1.AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN1.AGE_DATE_ORA(3) & sqle & ") AGE_3" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE                                                              <= " & TAC.ARCMAIN1.AGE_DATE_ORA(4) & sqle & ") AGE_4" & vbCrLf _
        & IIf(consider_future, _
          ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE > " & TAC.ARCMAIN1.DUE_DATE_ORA(1) & "                                                             " & sqle & ") DUE_0" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE > " & TAC.ARCMAIN1.DUE_DATE_ORA(2) & " AND ARTOPEN1.INV_DUE_DATE <= " & TAC.ARCMAIN1.DUE_DATE_ORA(1) & sqle & ") DUE_1" & vbCrLf, _
          ", 0 DUE_0" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN1.DUE_DATE_ORA(2) & "                                                            " & sqle & ") DUE_1" & vbCrLf) _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN1.DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & TAC.ARCMAIN1.DUE_DATE_ORA(2) & sqle & ") DUE_2" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN1.DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & TAC.ARCMAIN1.DUE_DATE_ORA(3) & sqle & ") DUE_3" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE                                                                   < " & TAC.ARCMAIN1.DUE_DATE_ORA(4) & sqle & ") DUE_4" & vbCrLf

        'AGED_TOTALS = "" _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & "'01-JAN-2012'" & "                                                         " & sqle & ") AGE_1" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & "'01-JAN-2012'" & " AND ARTOPEN1.INV_DATE <= " & "'01-JAN-2012'" & sqle & ") AGE_2" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & "'01-JAN-2012'" & " AND ARTOPEN1.INV_DATE <= " & "'01-JAN-2012'" & sqle & ") AGE_3" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DATE                                                              <= " & "'01-JAN-2012'" & sqle & ") AGE_4" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & "'01-JAN-2012'" & "                                                            " & sqle & ") DUE_1" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & "'01-JAN-2012'" & " AND ARTOPEN1.INV_DUE_DATE < " & "'01-JAN-2012'" & sqle & ") DUE_2" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & "'01-JAN-2012'" & " AND ARTOPEN1.INV_DUE_DATE < " & "'01-JAN-2012'" & sqle & ") DUE_3" & vbCrLf _
        ' & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE                                                                   < " & "'01-JAN-2012'" & sqle & ") DUE_4" & vbCrLf

    End Sub

    Public Shared Function Clean_Out_ARTPYMT2_Started_NOT_Completed(PYMT_BATCH_NO As String, PYMT_BATCH_LNO As Integer, Optional quiet As Boolean = True) As Boolean

        Dim ok_to_delete As Boolean = True

        For Each TABLE_NAME As String In New String() {"ARTPYMT3", "ARTPYMT4", "ARTPYMT5"}
            ASCMAIN1.sql = "Select Count (*) from " & TABLE_NAME _
                & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" _
                & "   and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
            Dim X As Integer = ASCDATA1.GetDataValue
            If X <> 0 Then
                If Not quiet Then
                    MsgBox("Orphan Rows in Application Only Batch-Line " & PYMT_BATCH_NO & "-" & CStr(PYMT_BATCH_LNO), _
                    MsgBoxStyle.OkOnly, "Please contact ABS")
                End If
                ok_to_delete = False
                Exit For
            End If
        Next

        If ok_to_delete Then
            ASCMAIN1.sql = "Delete from ARTPYMT2" _
                & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" _
                & "   and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Select Count (*) from ARTPYMT2" _
                & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            Dim X As Integer = Val(ASCDATA1.GetDataValue)
            If X = 0 Then
                ASCMAIN1.sql = "Delete from ARTPYMT1" _
                    & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
                ASCDATA1.ExecuteSQL()
            End If
        End If

        Return ok_to_delete
    End Function

    Public Shared Sub Create_ARTCBDA1(ByRef ARTCBDA1 As String, RYP0 As String, RYP1 As String)

        Dim INV_TYPEs As String = " IN ('B','D','C','O','I')" ' " = 'B'" ' 
        Dim FYP As String = ""
        If RYP0 <> "" Then FYP = ASCMAIN1.Period_Calc(RYP0, -1)

        Dim BAL As String = "" _
            & "Select GLTCREC3.DETL_CVX_NO CUST_CODE, NVL(ARTOPENX.REASON_CODE,ARTOPEN1.REASON_CODE) REASON_CODE" & vbCrLf _
            & ", 0 BEG_B, 0 NEW_B, 0 APP_B, 0 END_B, 0 BEG_C, 0 NEW_C, 0 APP_C, 0 END_C, 0 NEW_X" & vbCrLf _
            & " from GLTCREC3, ARTOPEN1, ARTOPENX, SOTINVH1" & vbCrLf _
            & " where GLTCREC3.CREC_TYPE_CODE = 'AR'" & vbCrLf _
            & "   and GLTCREC3.DETL_CTL_TYPE " & INV_TYPEs _
            & "   and GLTCREC3.OPS_YYYYPP = '000000'" & vbCrLf _
            & "   and ARTOPEN1.CUST_CODE (+) = GLTCREC3.DETL_CVX_NO" & vbCrLf _
            & "   and ARTOPEN1.INV_TYPE (+) = GLTCREC3.DETL_CTL_TYPE" & vbCrLf _
            & "   and ARTOPEN1.INV_NUM (+) = GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & "   and ARTOPENX.CUST_CODE (+) = GLTCREC3.DETL_CVX_NO" & vbCrLf _
            & "   and ARTOPENX.INV_TYPE (+) = GLTCREC3.DETL_CTL_TYPE" & vbCrLf _
            & "   and ARTOPENX.INV_NUM (+) = GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE (+) = GLTCREC3.DETL_CTL_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO (+) = GLTCREC3.DETL_CTL_NO" & vbCrLf _
            & IIf(ASCMAIN1.CLIENT = "INT", "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf, "") _
            & " group by GLTCREC3.DETL_CVX_NO, NVL(ARTOPENX.REASON_CODE,ARTOPEN1.REASON_CODE)"

        ASCMAIN1.sql = "" & vbCrLf _
            & Replace( _
              Replace( _
              Replace(BAL, "000000", FYP), _
              "0 BEG_B", "SUM (DECODE(GLTCREC3.DETL_CTL_TYPE,'B',GLTCREC3.CREC_AMT,'O',GLTCREC3.CREC_AMT,0)) BEG_B"), _
              "0 BEG_C", "SUM (DECODE(GLTCREC3.DETL_CTL_TYPE,'C',GLTCREC3.CREC_AMT,'D',GLTCREC3.CREC_AMT,'I',GLTCREC3.CREC_AMT,0)) BEG_C") _
            & " union " & vbCrLf _
            & "Select ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & ", 0 BEG_BAL" & vbCrLf _
            & ", SUM (DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'1',ARTPYMT5.GL_DIST_AMT,0)) NEW_B" & vbCrLf _
            & ", 0 APP_B, 0 END_B, 0 BEG_C, 0 NEW_C, 0 APP_C, 0 END_C" & vbCrLf _
            & ", SUM (DECODE(NVL(ARTPYMT5.CHARGEBACK_IND,'0'),'0',ARTPYMT5.GL_DIST_AMT,0)) NEW_X" & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ARTPYMT2.CUST_CODE, ARTPYMT3.REASON_CODE" & vbCrLf _
            & ", 0 BEG_B" & vbCrLf _
            & ", 0 NEW_B" & vbCrLf _
            & ", SUM (DECODE(ARTPYMT3.INV_TYPE,'B',NVL(ARTPYMT3.INV_PMT,0),'O',NVL(ARTPYMT3.INV_PMT,0),0)) APP_B" & vbCrLf _
            & ", 0 END_B" & vbCrLf _
            & ", 0 BEG_C" & vbCrLf _
            & ", 0 NEW_C" & vbCrLf _
            & ", SUM (DECODE(ARTPYMT3.INV_TYPE,'C',NVL(ARTPYMT3.INV_PMT,0),'D',NVL(ARTPYMT3.INV_PMT,0),'I',NVL(ARTPYMT3.INV_PMT,0),0)) APP_C" & vbCrLf _
            & ", 0 END_C" & vbCrLf _
            & ", 0 NEW_X" & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2,ARTPYMT3,SOTINVH1" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT3.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT3.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf _
            & "   and ARTPYMT3.INV_TYPE " & INV_TYPEs _
            & "   and SOTINVH1.INV_TYPE (+) = ARTPYMT3.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO (+) = ARTPYMT3.INV_NUM" & vbCrLf _
            & IIf(ASCMAIN1.CLIENT = "INT", "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf,"") _
            & " group by ARTPYMT2.CUST_CODE, ARTPYMT3.REASON_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE" & vbCrLf _
            & ", 0 BEG_B, 0 NEW_B, 0 APP_B, 0 END_B" & vbCrLf _
            & ", 0 BEG_C" & vbCrLf _
            & ", Sum(NVL(SOTINVH1.INV_TOTAL_AMOUNT,0)) NEW_C, 0 APP_C, 0 END_C" & vbCrLf _
            & ", 0 NEW_X" & vbCrLf _
            & " from SOTINVH1" & vbCrLf _
            & " where SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
            & "   and SOTINVH1.REASON_CODE is Not Null" _
            & "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf _
            & " group by SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & IIf(RYP1 = ASCMAIN1.CYP, "" & vbCrLf _
                    & "Select ARTOPEN1.CUST_CODE, ARTOPEN1.REASON_CODE" & vbCrLf _
                    & ", 0 BEG_B, 0 NEW_B, 0 APP_B, SUM (DECODE(ARTOPEN1.INV_TYPE,'B',ARTOPEN1.INV_BALANCE,'O',ARTOPEN1.INV_BALANCE,0)) END_B" & vbCrLf _
                    & ", 0 BEG_C, 0 NEW_C, 0 APP_C, SUM (DECODE(ARTOPEN1.INV_TYPE,'C',ARTOPEN1.INV_BALANCE,'D',ARTOPEN1.INV_BALANCE,'I',ARTOPEN1.INV_BALANCE,0)) END_C" & vbCrLf _
                    & ", 0 NEW_X FROM ARTOPEN1,SOTINVH1" & vbCrLf _
                    & " where ARTOPEN1.INV_TYPE IN ('B','D','C','O','I') AND ARTOPEN1.INV_BALANCE <> 0" & vbCrLf _
                    & "   and SOTINVH1.INV_TYPE (+) = ARTOPEN1.INV_TYPE" & vbCrLf _
                    & "   and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM" & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "INT", "   and (SOTINVH1.ORDR_TYPE_CODE = 'TOP' or SOTINVH1.ORDR_TYPE_CODE = 'DIF')" & vbCrLf, "") _
                    & " group by ARTOPEN1.CUST_CODE, ARTOPEN1.REASON_CODE", "" & vbCrLf _
                    & Replace( _
                      Replace( _
                      Replace(BAL, "000000", RYP1), _
                      "0 END_B", "SUM (DECODE(DETL_CTL_TYPE,'B',GLTCREC3.CREC_AMT,'O',GLTCREC3.CREC_AMT,0)) END_B"), _
                      "0 END_C", "SUM (DECODE(DETL_CTL_TYPE,'C',GLTCREC3.CREC_AMT,'D',GLTCREC3.CREC_AMT,0)) END_C"))


        If ARTCBDA1 = "" Then
            ARTCBDA1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Create Index I_" & ARTCBDA1 & "_1 on " & ARTCBDA1 & " (CUST_CODE, REASON_CODE)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ARTCBDA1 & " Add Primary Key (CUST_CODE,REASON_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Delete from " & ARTCBDA1)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDA1 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Public Shared Sub Create_ARTSTMT1(TABLE_NAME As String, CYPdt As Date)

        Dim sql_type As String = ""
        For Each T As String In New String() {"I", "R", "C", "D", "B", "O"}
            sql_type &= ", Sum (CASE WHEN INV_TYPE = '" & T & "' THEN INV_BALANCE ELSE 0 END) TYP_" & T & "_OPEN" & vbCrLf
            sql_type &= ", Sum (CASE WHEN INV_TYPE = '" & T & "' THEN 1 ELSE 0 END) TYP_" & T & "_CNT" & vbCrLf
        Next

        Dim sql_ARTOPEN1 = "Select ARTCUST1.CUST_CODE CUST_BILL_TO_CUST" & vbCrLf _
            & ", SUM(ARTOPEN1.INV_BALANCE) INV_BALANCE " & vbCrLf _
            & TAC.ARCMAIN1.AGED_TOTALS _
            & sql_type _
            & " from ARTOPEN1,ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE " & vbCrLf _
            & " and ARTOPEN1.INV_BALANCE <> 0" & vbCrLf _
            & " and NVL(ARTOPEN1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "') <= '" & ASCMAIN1.CYP & "'" _
            & " group by ARTCUST1.CUST_CODE"

        ASCMAIN1.sql = "Insert into " & TABLE_NAME & " ARTSTMT1 (OPS_YYYYPP,CUST_CODE" _
            & ",AGE_0,AGE_1,AGE_2,AGE_3,AGE_4" _
            & ",DUE_0,DUE_1,DUE_2,DUE_3,DUE_4" _
            & ",TYP_I_OPEN,TYP_R_OPEN,TYP_C_OPEN,TYP_D_OPEN,TYP_B_OPEN,TYP_O_OPEN" _
            & ",TYP_I_CNT,TYP_R_CNT,TYP_C_CNT,TYP_D_CNT,TYP_B_CNT,TYP_O_CNT) " _
            & " Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, CUST_CODE" _
            & ",AGE_0,AGE_1,AGE_2,AGE_3,AGE_4" _
            & ",DUE_0,DUE_1,DUE_2,DUE_3,DUE_4" _
            & ",TYP_I_OPEN,TYP_R_OPEN,TYP_C_OPEN,TYP_D_OPEN,TYP_B_OPEN,TYP_O_OPEN" _
            & ",TYP_I_CNT,TYP_R_CNT,TYP_C_CNT,TYP_D_CNT,TYP_B_CNT,TYP_O_CNT" _
            & " from (Select CUST_CODE, X.*" _
            & " from ARTCUST1,(" & sql_ARTOPEN1 & ") X where X.CUST_BILL_TO_CUST (+) = ARTCUST1.CUST_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & TABLE_NAME & " ARTSTMT1 Set TOTAL_DUE = NVL(DUE_1,0)+NVL(DUE_2,0)+NVL(DUE_3,0)+NVL(DUE_4,0) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & TABLE_NAME & " ARTSTMT1 Set TOTAL_OPEN_AMT = NVL(AGE_0,0)+NVL(AGE_1,0)+NVL(AGE_2,0)+NVL(AGE_3,0)+NVL(AGE_4,0) where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select CUST_CODE, TOTAL_DUE from " & TABLE_NAME & " ARTSTMT1" & vbCrLf _
            & "   where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & TABLE_NAME & " ARTSTMT1" & vbCrLf _
            & "    Set BALFWD = R1.TOTAL_DUE where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and CUST_CODE= R1.CUST_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into " & TABLE_NAME & " ARTSTMT1 (CUST_CODE, OPS_YYYYPP, BALFWD) " & vbCrLf _
            & "     values (R1.CUST_CODE,'" & ASCMAIN1.CYP & "',R1.TOTAL_DUE);" & vbCrLf _
            & "   End If;" _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        'Dim sql_act As String = "" _
        '    & "Select ARTCUST1.CUST_CODE CUST_BILL_TO_CUST, ARTCUST1.CUST_CREDIT_GROUP_CUST, " & vbCrLf _
        '    & " ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, " & vbCrLf _
        '    & " ARTPYMT2.CUST_PYMT_AMT, ARTPYMT1.PYMT_BATCH_DATE," & vbCrLf _
        '    & " SUM (ARTPYMT3.INV_PMT) AR_DOLLARS," & vbCrLf _
        '    & " SUM (ARTPYMT3.INV_PMT * (ARTPYMT1.PYMT_BATCH_DATE - ARTPYMT3.INV_DATE)) DAY_DOLLARS," & vbCrLf _
        '    & " COUNT (DECODE (ARTPYMT3.INV_TYPE, 'I', 1, 0)) NO_INVOICES," & vbCrLf _
        '    & " SUM (DECODE (ARTPYMT3.INV_TYPE, 'O', ARTPYMT3.INV_PMT, 0)) DOLLARS_PAID_OA" & vbCrLf _
        '    & " From ARTPYMT1, ARTPYMT2, ARTPYMT3, ARTCUST1" & vbCrLf _
        '    & " WHERE ARTPYMT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
        '    & " AND ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
        '    & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
        '    & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
        '    & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
        '    & " GROUP BY ARTCUST1.CUST_CODE, ARTCUST1.CUST_CREDIT_GROUP_CUST," & vbCrLf _
        '    & " ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO," & vbCrLf _
        '    & " ARTPYMT2.CUST_PYMT_AMT , ARTPYMT1.PYMT_BATCH_DATE"

        'ASCMAIN1.sql = "" _
        '    & "Begin " _
        '    & " Declare Cursor C1 is " _
        '    & "  Select CUST_BILL_TO_CUST CUST_CODE" _
        '    & ",  SUM (CUST_PYMT_AMT) CUST_PYMT_AMT" _
        '    & ",  SUM (DOLLARS_PAID_OA) DOLLARS_PAID_OA" _
        '    & ",  SUM (DAY_DOLLARS) DAY_DOLLARS" _
        '    & ",  SUM (NO_INVOICES) NO_INVOICES" _
        '    & "   from (" & sql_act & ") X group by CUST_BILL_TO_CUST;" _
        '    & " Begin" _
        '    & "  For R1 in C1 Loop" _
        '    & "   Update ARTSTMT1 Set " _
        '    & "    DAY_DOLLARS = R1.DAY_DOLLARS," _
        '    & "    DOLLARS_PAID = R1.CUST_PYMT_AMT + R1.DOLLARS_PAID_OA," _
        '    & "    NO_INVOICES = R1.NO_INVOICES" _
        '    & "    Where CUST_CODE = R1.CUST_CODE;" _
        '    & "  End Loop;" _
        '    & " End;" _
        '    & "End;"
        'ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select * from ARTCUST6;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & TABLE_NAME & " ARTSTMT1 Set " & vbCrLf _
            & "    CUST_HIGH_BAL_DATE = R1.CUST_HIGH_BAL_DATE," & vbCrLf _
            & "    CUST_HIGH_BAL_AMT = R1.CUST_HIGH_BAL_AMT" & vbCrLf _
            & "    Where CUST_CODE = R1.CUST_CODE" & vbCrLf _
            & "      and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub
End Class