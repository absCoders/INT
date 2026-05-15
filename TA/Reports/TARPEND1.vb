Public Class TARPEND1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)

        ASCMAIN1.sql = "SELECT ASTOPST1.MENU_ITEM_OBJECT, ASTOPST1.INIT_DATE" _
            & ", ASTOPST1.USER_ID, ASTMENU1.MENU_ITEM_DESC" _
            & " FROM ASTOPST1,ASTMENU1 " _
            & " WHERE ASTOPST1.YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " AND ASTOPST1.PRD_CLOSE_IND = '1' AND ASTOPST1.UPDATED = '1'" _
            & " AND ASTMENU1.MENU_ID = ASTOPST1.MENU_ID" _
            & " AND ASTMENU1.MENU_ITEM_OBJECT <> 'TARPEND1'" _
            & " AND ASTMENU1.MENU_ITEM_TYPE = ASTOPST1.MENU_ITEM_TYPE" _
            & " AND ASTMENU1.MENU_ITEM_OBJECT = ASTOPST1.MENU_ITEM_OBJECT"
        Dim sql As String = ASCMAIN1.sql

        Dim tblTATPEND1 As DataTable = ASCDATA1.GetDataTable("", "TATPEND1")

        With dst
            ASCMAIN1.sql = "Select SOTPRIC2.* from SOTPRIC2"
            Create_TDA(.Tables.Add, "SOTPRIC2", "**", 0, True, "", 2)
        End With


        grdTATPEND1.DataSource = tblTATPEND1
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "U"

    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" <> "1" Then
                EMsg = EMsg & vbCr & "Period-End has not been Initialized"
            End If
            Dim z As String = Absx1.cmbFor("RYP").Text
            z = Mid(z, 1, 4) & Mid(z, 6, 2)
            Dim zctl As String = ASCDATA1.GetDataValue("Select CURR_YEAR || CURR_PERIOD from ASTPCTL1") & ""
            If zctl <> z Then
                EMsg = EMsg & vbCr & "Incorrect Period to Finalize"
            End If

            If EMsg = "" Then
                Check_for_Records("ICTIADJ1", "Inventory Adjustment Journal", "NVL(JOURNAL_IND,'0') = '0'")
                Check_for_Records("ICTIXFR1", "Warehouse Transfer Journal", "NVL(JOURNAL_IND,'0') = '0'")
                Check_for_Records("ICTIREC1", "PO Receipts Journal", "NVL(JOURNAL_IND,'0') = '0'")

                If EMsg <> "" Then
                    EMsg = "Cannot Proceed because a Clean Cut-off has not been established as follows:" & vbCr & EMsg
                End If
            End If
        End If

    End Sub

    Sub Check_for_Records( _
    ByVal TABLE_NAME As String, _
    ByVal TABLE_DESC As String, _
    Optional ByVal where_clause As String = "", _
    Optional ByVal custom_sql As String = "")

        If custom_sql <> "" Then
            ASCMAIN1.sql = custom_sql
        Else
            ASCMAIN1.sql = "Select count (*) from " & TABLE_NAME
            If where_clause <> "" Then
                ASCMAIN1.sql &= " where " & where_clause
            End If
        End If

        Dim sql As String = ASCMAIN1.sql
        Dim r As Long = Val(ASCDATA1.GetDataValue() & "")
        If r <> 0 Then
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(r) & " Records"
        End If

    End Sub

    Overrides Sub Update_Record()

        Get_PARM("ICTPARM1")

        Dim NYP As String = ASCMAIN1.CYP
        If Mid$(NYP, 5, 2) = "12" Then
            Dim YYYY As Integer = Val(Mid$(NYP, 1, 4))
            Mid$(NYP, 5, 2) = "01"
            Mid$(NYP, 1, 4) = Format$(YYYY + 1, "0000")
        Else
            Dim PP As Integer = Val(Mid$(NYP, 5, 2))
            Mid$(NYP, 5, 2) = Format$(PP + 1, "00")
        End If

        ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        Dim CYPdt As Date = CDate(ASCDATA1.GetDataValue) '.AddDays(1)

        ' A/R

        Get_PARM("ARTPARM1")
        TAC.ARCMAIN1.Get_Aging_Data(ROWs("ARTPARM1"), CYPdt, True, True)
        TAC.ARCMAIN1.Create_ARTSTMT1("ARTSTMT1", CYPdt)

       

        ASCMAIN1.Progress("A/R Closed Items Purge", "")
        ASCDATA1.ExecuteSQL("Update ARTOPEN1 set OPS_YYYYPP_F = '" & ASCMAIN1.CYP & "' where INV_BALANCE = 0 ")
        ASCDATA1.ExecuteSQL("Delete from ARTOPENX where (CUST_CODE, INV_TYPE, INV_NUM) in (Select CUST_CODE, INV_TYPE, INV_NUM from ARTOPEN1 where OPS_YYYYPP_F is Not Null)")
        ASCDATA1.ExecuteSQL("Insert into ARTOPENX Select * from ARTOPEN1 where OPS_YYYYPP_F is Not Null")
        ASCDATA1.ExecuteSQL("Delete from ARTOPEN1 where OPS_YYYYPP_F is Not Null")


        ' Sales Rep Snapshots
        ASCMAIN1.Progress("Sales Rep Snapshots", "")
        ASCMAIN1.sql = "Insert into ARTCUST4 Select '" & RYP & "' OPS_YYYYPP, ARTCUST1.CUST_CODE" _
            & ", ARTCUST1.SREP_CODE, ARTCUST1.SREP2_CODE, ARTCUST1.SREP_CODE_OVER" _
            & ", DECODE(SOTSREP2.SREP_CODE,NULL,SOTSREP1.SREP_COMM_PCT,SOTSREP2.SREP_COMM_PCT) SREP_COMM_PCT" _
            & ", DECODE(SOTSREP2.SREP_CODE,NULL,SOTSREP1.SREP_COMM_PCT_SPEC,SOTSREP2.SREP_COMM_PCT_SPEC) SREP_COMM_PCT_SPEC" _
            & ", DECODE(SOTSREP2_2.SREP_CODE,NULL,SOTSREP1_2.SREP_COMM_PCT,SOTSREP2_2.SREP_COMM_PCT) SREP2_COMM_PCT" _
            & ", DECODE(SOTSREP2_2.SREP_CODE,NULL,SOTSREP1_2.SREP_COMM_PCT_SPEC,SOTSREP2_2.SREP_COMM_PCT_SPEC) SREP2_COMM_PCT_SPEC" _
            & ", ARTCUST1.SREP_COMM_PCT_OVER" _
            & " from ARTCUST1,SOTSREP1 SOTSREP1, SOTSREP1 SOTSREP1_2, SOTSREP2, SOTSREP2 SOTSREP2_2" _
            & " where SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE" _
            & "   and SOTSREP1_2.SREP_CODE (+) = ARTCUST1.SREP2_CODE" _
            & "   and SOTSREP2.SREP_CODE (+) = ARTCUST1.SREP_CODE" _
            & "   and SOTSREP2.CUST_CODE (+) = ARTCUST1.CUST_CODE" _
            & "   and SOTSREP2_2.SREP_CODE (+) = ARTCUST1.SREP2_CODE" _
            & "   and SOTSREP2_2.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ARTCUST7 Select '" & RYP & "' OPS_YYYYPP" _
            & ", ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" _
            & ", ARTCUST2.SREP_CODE, ARTCUST2.SELL_CODE" _
            & ", ARTCUST2.SELL_CODE_AC" _
            & " from ARTCUST2"
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.CLIENT = "INT" Then
            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO from ARTCUST2, (" & vbCrLf _
                & "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", SUM (CASE WHEN OPS_YYYYPP_OPENED IS NOT NULL" & vbCrLf _
                & "  and (OPS_YYYYPP_CLOSED IS NULL OR OPS_YYYYPP_CLOSED > '" & NYP & "') THEN 1 ELSE 0 END) OPENS" & vbCrLf _
                & " from SATAUTH1" & vbCrLf _
                & " group by CUST_CODE, CUST_STORE_NO) X" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = X.CUST_CODE and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO and ARTCUST2.CUST_STORE_STATUS = 'A' and X.OPENS = 0"
            ASCDATA1.ExecuteSQL("Insert into ASTAUDT1 Select 'ARTCUST2', CUST_CODE || ':' || CUST_STORE_NO, 'CUST_STORE_STATUS'" & vbCrLf _
                                & ",'" & ASCMAIN1.USER_ID & "', SYSDATE, 'A', 'I', 'E'" & vbCrLf _
                                & ",'No High Collections Open', NULL, NULL" & vbCrLf _
                                & ", '" & ASCMAIN1.SESSION_NO & "', " & CStr(SELECTION_NO) & ",'" & XNO & "'" & vbCrLf _
                                & " from (" & ASCMAIN1.sql & ")")

            ' ASCDATA1.ExecuteSQL("Update ARTCUST2 Set CUST_STORE_STATUS = 'I' where (CUST_CODE, CUST_STORE_NO) in (" & ASCMAIN1.sql & ")")
            ' note - LBM wants to disable this for now 05/08/2018
            ' RE-ENABLING SINCE WE ARE SAVING THE LAST AE - SEE SP APPROVAL EMAIL 5/31 RE: removal of AE and RSC when door is closed
            ASCDATA1.ExecuteSQL("Update ARTCUST2 Set SELL_CODE_LAST = SELL_CODE, SELL_CODE_LAST_YP = '" & ASCMAIN1.CYP & "' where SELL_CODE IS NOT NULL AND (CUST_CODE, CUST_STORE_NO) in (" & ASCMAIN1.sql & ")")
            ASCDATA1.ExecuteSQL("Update ARTCUST2 Set SELL_CODE_AC_LAST = SELL_CODE_AC, SELL_CODE_AC_LAST_YP = '" & ASCMAIN1.CYP & "' where SELL_CODE_AC IS NOT NULL AND (CUST_CODE, CUST_STORE_NO) in (" & ASCMAIN1.sql & ")")
            ASCDATA1.ExecuteSQL("Update ARTCUST2 Set CUST_STORE_STATUS = 'I', SDS_CODE = NULL, SELL_CODE = NULL, SELL_CODE_AC = NULL where (CUST_CODE, CUST_STORE_NO) in (" & ASCMAIN1.sql & ")")
        End If

        ASCMAIN1.Progress("Reset Customer MTD Activity Summary", "")
        ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
            & " CUST_SALES_MTD = 0" & vbCrLf _
            & ", CUST_COGS_MTD = 0" & vbCrLf _
            & ", CUST_CASH_MTD = 0" & vbCrLf _
            & ", CUST_FIN_CHG_MTD = 0" & vbCrLf _
            & ", CUST_NUM_INV_MTD = 0" & vbCrLf _
            & ", CUST_NUM_FIN_MTD = 0" & vbCrLf _
            & ", CUST_CRED_MTD = 0" & vbCrLf _
            & ", CUST_HIGH_BAL_DATE = NULL" & vbCrLf _
            & ", CUST_HIGH_BAL_AMT = 0"
        ASCDATA1.ExecuteSQL()

        If Mid$(ASCMAIN1.CYP, 5, 2) = "12" Then
            ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
                & "  CUST_SALES_LYR = CUST_SALES_YTD" & vbCrLf _
                & ", CUST_COGS_LYR = CUST_COGS_YTD" & vbCrLf _
                & ", CUST_CASH_LYR = CUST_CASH_YTD" & vbCrLf _
                & ", CUST_FIN_CHG_LYR = CUST_FIN_CHG_YTD" & vbCrLf _
                & ", CUST_NUM_INV_LYR = CUST_NUM_INV_YTD" & vbCrLf _
                & ", CUST_NUM_FIN_LYR = CUST_NUM_FIN_YTD" & vbCrLf _
                & ", CUST_CRED_LYR = CUST_CRED_YTD"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update ARTCUST6 Set" & vbCrLf _
                & "  CUST_SALES_YTD = 0" & vbCrLf _
                & ", CUST_COGS_YTD = 0" & vbCrLf _
                & ", CUST_CASH_YTD = 0" & vbCrLf _
                & ", CUST_FIN_CHG_YTD = 0" & vbCrLf _
                & ", CUST_NUM_INV_YTD = 0" & vbCrLf _
                & ", CUST_NUM_FIN_YTD = 0" & vbCrLf _
                & ", CUST_CRED_YTD = 0"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select CUST_CODE, SUM (INV_BALANCE) INV_BALANCE" & vbCrLf _
            & " from ARTOPEN1 group by CUST_CODE;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ARTCUST6 Set CUST_HIGH_BAL_DATE = TRUNC(SYSDATE)," & vbCrLf _
            & " CUST_HIGH_BAL_AMT = R1.INV_BALANCE" & vbCrLf _
            & " where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  INSERT INTO ARTCUST6 (CUST_CODE, CUST_HIGH_BAL_DATE, CUST_HIGH_BAL_AMT)" & vbCrLf _
            & "  VALUES (R1.CUST_CODE, TRUNC(SYSDATE), R1.INV_BALANCE);" & vbCrLf _
            & " End If;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()




        ' A/P
        ASCMAIN1.Progress("Reset Vendor MTD Activity Summary", "")

        ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
            & "  VEND_PURCHASES_MTD = 0" & vbCrLf _
            & ", VEND_PAYMENTS_MTD = 0" & vbCrLf _
            & ", VEND_DISC_TAKEN_MTD = 0" & vbCrLf _
            & ", VEND_NUM_INV_MTD = 0" & vbCrLf _
            & ", VEND_NUM_CHKS_MTD = 0" & vbCrLf
        ASCDATA1.ExecuteSQL()

        If Mid$(ASCMAIN1.CYP, 5, 2) = "12" Then
            ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
                & "  VEND_PURCHASES_LYR = VEND_PURCHASES_YTD" & vbCrLf _
                & ", VEND_PAYMENTS_LYR = VEND_PAYMENTS_YTD" & vbCrLf _
                & ", VEND_DISC_TAKEN_LYR = VEND_DISC_TAKEN_YTD" & vbCrLf _
                & ", VEND_NUM_INV_LYR = VEND_NUM_INV_YTD" & vbCrLf _
                & ", VEND_NUM_CHKS_LYR = VEND_NUM_CHKS_YTD"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update APTVEND5 Set" & vbCrLf _
                & "  VEND_PURCHASES_YTD = 0" & vbCrLf _
                & ", VEND_PAYMENTS_YTD = 0" & vbCrLf _
                & ", VEND_DISC_TAKEN_YTD = 0" & vbCrLf _
                & ", VEND_NUM_INV_YTD = 0" & vbCrLf _
                & ", VEND_NUM_CHKS_YTD = 0"
            ASCDATA1.ExecuteSQL()
        End If

        ' Purge Files
        Purge_Files()

        ' Forecasts

        ASCMAIN1.sql = "" _
            & "Insert into DPTITMF1" & vbCrLf _
            & " Select '" & NYP & "' OPS_YYYYPP" & vbCrLf _
            & ", ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC, FORECAST" & vbCrLf _
            & " from DPTITMF1" & vbCrLf _
            & " where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and OPS_YYYYPP_FC >= '" & NYP & "'"
        ASCDATA1.ExecuteSQL()

        Create_TDA(dst.Tables.Add, "DPTITMF1", "*")

        Dim sqlMARKET_CODEs_with_PD As String = "" _
            & "Select MARKET_CODE from SOTMKTC1" & vbCrLf _
            & " where (NVL(PAST_DUE_FC_SB,'0') <> '0'" & vbCrLf _
            & "     or NVL(PAST_DUE_FC_NB,'0') <> '0'" & vbCrLf _
            & "     or NVL(PAST_DUE_FC_UB,'0') <> '0'" & vbCrLf _
            & "     or NVL(PAST_DUE_FC_SP,'0') <> '0'" & vbCrLf _
            & "     or NVL(PAST_DUE_FC_NP,'0') <> '0'" & vbCrLf _
            & "     or NVL(PAST_DUE_FC_UP,'0') <> '0')"

        ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" & vbCrLf _
            & ", SUM (FORECAST) FORECAST" & vbCrLf _
            & ", SUM (QTY_SHIPPED) QTY_SHIPPED, SUM (QTY_RETURNED) QTY_RETURNED from " & vbCrLf _
            & " (Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE ITEM_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", SUM (0) FORECAST, SUM (DECODE (INV_TYPE, 'I', ORDR_QTY_SHIP, 0)) QTY_SHIPPED" & vbCrLf _
            & ", SUM (DECODE (INV_TYPE, 'C', ORDR_QTY_SHIP, 0)) QTY_RETURNED" & vbCrLf _
            & " from SOTINVH2, ARTCUST1, SOTTCLS1, ICTITEM1,SOTMKTC1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "' and ORDR_QTY_SHIP <> 0 " & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE " & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE " & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & "   and NVL(SOTMKTC1.MARKET_CODE,SOTTCLS1.MARKET_CODE) in (" & sqlMARKET_CODEs_with_PD & ")" & vbCrLf _
            & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?')))" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO " & vbCrLf _
            & " UNION " & vbCrLf _
            & "Select DPTITMF1.MARKET_CODE, DPTITMF1.ITEM_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO, " & vbCrLf _
            & " SUM (DPTITMF1.FORECAST) FORECAST, SUM (0) QTY_SHIPPED, SUM (0) QTY_RETURNED" & vbCrLf _
            & " from DPTITMF1, ICTITEM1 where DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " and (DPTITMF1.OPS_YYYYPP_FC = DPTITMF1.OPS_YYYYPP or DPTITMF1.OPS_YYYYPP_FC = '000000')" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = DPTITMF1.ITEM_CODE and DPTITMF1.MARKET_CODE in (" & sqlMARKET_CODEs_with_PD & ")" & vbCrLf _
            & " and DPTITMF1.FORECAST <> 0" & vbCrLf _
            & " group by DPTITMF1.MARKET_CODE, DPTITMF1.ITEM_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO )" & vbCrLf _
            & " GROUP BY MARKET_CODE, ITEM_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" & vbCrLf

        Dim MARKET_CODE As String = ""
        Dim rowSOTMKTC1 As DataRow = Nothing

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "MARKET_CODE, ITEM_CODE, ITEM_SNU_CODE")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            If MARKET_CODE <> row.Item("MARKET_CODE") Then
                MARKET_CODE = row.Item("MARKET_CODE")
                rowSOTMKTC1 = LookUp("SOTMKTC1", MARKET_CODE)
            End If
            Dim QTY_SHIPPED As Int64 = Val(row.Item("QTY_SHIPPED") & "")
            Dim QTY_RETURNED As Int64 = Val(row.Item("QTY_RETURNED") & "")
            Dim FORECAST As Int64 = Val(row.Item("FORECAST") & "")
            Dim PD_FORECAST As Int64 = 0

            'f PD_FORECAST <> 0 Then
            If rowSOTMKTC1.Item("PAST_DUE_FC_GRS_NET") & "" = "N" Then
                PD_FORECAST = FORECAST - (QTY_SHIPPED + QTY_RETURNED)
            Else
                PD_FORECAST = FORECAST - QTY_SHIPPED
            End If
            'End If

            Dim ITEM_SNU_CODE As String = row.Item("ITEM_SNU_CODE") & ""
            Dim ITEM_BASIC_PROMO As String = row.Item("ITEM_BASIC_PROMO") & ""
            If PD_FORECAST <> 0 Then
                Dim SNU_BP As String = ITEM_SNU_CODE & ITEM_BASIC_PROMO
                If Len(SNU_BP) = 2 And InStr("SNU", Mid$(SNU_BP, 1, 1)) <> 0 And InStr("BP", Mid$(SNU_BP, 2, 1)) <> 0 Then
                    Dim PD_IND As String = rowSOTMKTC1.Item("PAST_DUE_FC_" & SNU_BP) & ""
                    If PD_IND = "" Then PD_IND = "0"

                    ' NOTE THAT AT CCU THIS FIELD WAS MORE THAN JUST A BINARY
                    ' A          Pos or Neg PD FC                                            
                    ' P          Pos PD FC Only                                              
                    ' N          Neg PD FC Only                                              
                    ' 0          Zero PD FC 

                    If PD_IND = "0" Or (PD_FORECAST > 0 And PD_IND = "N") Or (PD_FORECAST < 0 And PD_IND = "P") Then
                        PD_FORECAST = 0
                    End If
                Else
                    PD_FORECAST = 0
                End If
            End If

            If PD_FORECAST <> 0 Then
                Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                With rowDPTITMF1
                    .Item("OPS_YYYYPP") = NYP
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("MARKET_CODE") = MARKET_CODE
                    .Item("OPS_YYYYPP_FC") = "000000"
                    .Item("FORECAST") = PD_FORECAST
                End With
                dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
            End If
        Next

        Update_Record_TDA("DPTITMF1")


        ' 12 Month Rolling Forecast


        Dim CYP12 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select SOTINVH2.ITEM_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_PLAN_QUIET_ZONE_TYPE, SUM (SOTINVH2.ORDR_QTY_SHIP) FORECAST" & vbCrLf _
            & "   from SOTINVH2,ARTCUST1,SOTTCLS1,SOTMKTC1,ICTITEM1" & vbCrLf _
            & "   where SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "     and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "     and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "     and SOTMKTC1.MARKET_CODE = SOTTCLS1.MARKET_CODE" & vbCrLf _
            & "     and SOTMKTC1.MARKET_AUTO_FC = '1'" & vbCrLf _
            & "     and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "     and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "     and SOTINVH2.ITEM_CODE in (Select ITEM_CODE from ICTITEM1 where PROD_CODE = 'SB')" & vbCrLf _
            & "   group by SOTINVH2.ITEM_CODE, SOTTCLS1.MARKET_CODE, ICTITEM1.ITEM_PLAN_QUIET_ZONE_TYPE" & vbCrLf _
            & "   having SUM(SOTINVH2.ORDR_QTY_SHIP) <> 0;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  Delete from DPTITMF1" & vbCrLf _
            & "   where OPS_YYYYPP = '" & NYP & "' and OPS_YYYYPP_FC = '" & CYP12 & "'" & vbCrLf _
            & "     and MARKET_CODE in (Select MARKET_CODE from SOTMKTC1 where MARKET_AUTO_FC = '1')" & vbCrLf _
            & "     and ITEM_CODE in (Select ITEM_CODE from ICTITEM1 where PROD_CODE = 'SB');" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   If NVL(R1.ITEM_PLAN_QUIET_ZONE_TYPE,'?') <> '4' Then " & vbCrLf _
            & "    Insert into DPTITMF1 Values ('" & NYP & "',R1.ITEM_CODE,R1.MARKET_CODE,'" & CYP12 & "',R1.FORECAST);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        If Mid(ASCMAIN1.CYM, 5, 2) = "01" Or Mid(ASCMAIN1.CYM, 5, 2) = "07" Then
            Dim OPS_YYYY As String = Mid(ASCMAIN1.CYM, 1, 4)
            Dim SEASON As String = IIf(Mid(ASCMAIN1.CYM, 5, 2) = "01", "S", "F")

            ' ABC Info

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select DPTABCP0.ITEM_CODE, DPTABCP0.ITEM_ABC_CODE" & vbCrLf _
                & ", DPTABCP2.ABC_MAX_POS, DPTABCP2.ABC_MIN_POS, DPTABCP2.ABC_MIN_DAYS_SUPPLY" & vbCrLf _
                & " from DPTABCP0, DPTABCP2" & vbCrLf _
                & " where DPTABCP2.ABC_CODE (+) = DPTABCP0.ITEM_ABC_CODE" & vbCrLf _
                & "   and DPTABCP2.ITEM_CATGY_CODE (+) = DPTABCP0.ITEM_CATGY_CODE" & vbCrLf _
                & "   and DPTABCP0.OPS_YYYY = '" & OPS_YYYY & "'" & vbCrLf _
                & "   and DPTABCP0.SEASON = '" & SEASON & "';" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & "  Update ICTITEM1 Set ITEM_ABC_CODE = R1.ITEM_ABC_CODE" & vbCrLf _
                & ", ITEM_POS_MAX = R1.ABC_MAX_POS" & vbCrLf _
                & ", ITEM_POS_MIN = R1.ABC_MIN_POS" & vbCrLf _
                & ", ITEM_MIN_DAYS_SUPPLY = R1.ABC_MIN_DAYS_SUPPLY" & vbCrLf _
                & "   where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            ' ITEM_CATGY_CODE

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " & vbCrLf _
                & " Select DPTPROJ0.ITEM_CODE, DPTPROJ0.ITEM_CATGY_CODE" & vbCrLf _
                & " from DPTPROJ0" & vbCrLf _
                & " where DPTPROJ0.OPS_YYYY = '" & OPS_YYYY & "'" & vbCrLf _
                & "   and DPTPROJ0.SEASON = '" & SEASON & "';" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & "  Update ICTITEM1 " & vbCrLf _
                & "   Set ITEM_CATGY_CODE = R1.ITEM_CATGY_CODE" & vbCrLf _
                & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If

        ' Inventory Statistics recorded with NYP
        ASCMAIN1.Progress("Inventory Statistics for Next Period", "")
        ASCMAIN1.sql = "Update ICTSTAT1 " & vbCrLf _
            & " Set OPS_YYYYPP = '000000' where OPS_YYYYPP = '" & NYP & "'"
        ASCDATA1.ExecuteSQL()

        ' Inventory Statistics
        ASCMAIN1.Progress("Inventory Statistics", "")
        ASCMAIN1.sql = "Insert into ICTSTAT1 " & vbCrLf _
            & " (OPS_YYYYPP, ITEM_CODE, WHSE_CODE, WHSE_QTY_BEG) " & vbCrLf _
            & " Select '" & NYP & "'" & vbCrLf _
            & " OPS_YYYYPP, ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND" & vbCrLf _
            & " from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0"
        ASCDATA1.ExecuteSQL()

        ' Inventory Status History
        ASCMAIN1.Progress("Inventory Status History", "")
        ASCMAIN1.sql = "Insert into ICTSTAT5 Select '" & ASCMAIN1.CYP & "',ICTSTAT2.* from ICTSTAT2 where " & vbCrLf _
            & " NVL(WHSE_QTY_ON_HAND,0) <> 0 or NVL(WHSE_QTY_ONPO,0) <> 0 or NVL(WHSE_QTY_PLAN,0) <> 0 or " & vbCrLf _
            & " NVL(WHSE_QTY_OPEN,0) <> 0 or NVL(WHSE_QTY_PICK,0) <> 0 or " & vbCrLf _
            & " NVL(WHSE_QTY_HOLD,0) <> 0 or NVL(WHSE_QTY_COMM,0) <> 0"
        ASCDATA1.ExecuteSQL()

        ' Correct Beg/End using Inventory Statistics recorded with NYP
        ASCMAIN1.Progress("Correct Beg/End Invty using Statistics for Next Period", "")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from ICTSTAT1 where OPS_YYYYPP = '000000' for Update;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT1" & vbCrLf _
            & "    Set WHSE_QTY_BEG = NVL(WHSE_QTY_BEG,0) + NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0)" & vbCrLf _
            & "      , WHSE_QTY_SHP = R1.WHSE_QTY_SHP, WHSE_QTY_RTN = R1.WHSE_QTY_RTN" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & NYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then " & vbCrLf _
            & "    Insert into ICTSTAT1 (OPS_YYYYPP,ITEM_CODE,WHSE_CODE,WHSE_QTY_BEG,WHSE_QTY_SHP,WHSE_QTY_RTN)" & vbCrLf _
            & "     Values ('" & NYP & "',R1.ITEM_CODE,R1.WHSE_CODE,NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0),R1.WHSE_QTY_SHP,R1.WHSE_QTY_RTN);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "   Update ICTSTAT5" & vbCrLf _
            & "    Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0)" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then " & vbCrLf _
            & "    Insert into ICTSTAT5 (OPS_YYYYPP,ITEM_CODE,WHSE_CODE,WHSE_QTY_ON_HAND)" & vbCrLf _
            & "     Values ('" & ASCMAIN1.CYP & "',R1.ITEM_CODE,R1.WHSE_CODE,NVL(R1.WHSE_QTY_SHP,0) - NVL(R1.WHSE_QTY_RTN,0));" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "   Delete from ICTSTAT1 where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' Cost History
        ASCMAIN1.Progress("Cost History", "")
        ASCMAIN1.sql = "Insert into ICTCOSTA" _
        & " Select '" & NYP & "', ICTCOSTC.* from ICTCOSTC"
        ASCDATA1.ExecuteSQL()

        ' Retail Price Queued Changes
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select ITEM_CODE, ITEM_RETAIL_PRICE, ITEM_NEW_RETAIL_PRICE, ITEM_NEW_RETAIL_PRICE_DATE" & vbCrLf _
            & "   from ICTITEM1" _
            & "   where ITEM_NEW_RETAIL_PRICE_DATE IS NOT NULL " _
            & "     and ITEM_NEW_RETAIL_PRICE_DATE <= '" & Format(CYPdt.AddDays(1), "dd-MMM-yyyy") & "';" _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" _
            & "   Update ICTITEM1 Set" _
            & "     ITEM_RETAIL_PRICE = R1.ITEM_NEW_RETAIL_PRICE," _
            & "     ITEM_NEW_RETAIL_PRICE_DATE = NULL," _
            & "     ITEM_NEW_RETAIL_PRICE = NULL" _
            & "    where ITEM_CODE = R1.ITEM_CODE;" _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' Retail Price History
        ASCMAIN1.Progress("Price History", "")
        ASCMAIN1.sql = "Insert into ICTRETLA Select ITEM_CODE" & vbCrLf _
            & ", '" & NYP & "' OPS_YYYYPP, ITEM_RETAIL_PRICE, ITEM_PRICE, ITEM_CATGY_CODE" & vbCrLf _
            & " from ICTITEM1"
        ASCDATA1.ExecuteSQL()



        ' Price List Changes

        With dst
            ASCMAIN1.sql = "Select SOTPRIC2.* from SOTPRIC2"
            Create_TDA(.Tables.Add, "SOTPRIC2", "**", 0, True, "", 2)
        End With

        ASCMAIN1.Progress("Price List Changes", "")
        ASCMAIN1.sql = "Select * from SOTPRIC2" & vbCrLf _
            & " where ITEM_NEW_PRICE_DATE <= '" & Format(CYPdt.AddDays(1), "dd-MMM-yyyy") & "'"
        Fill_Records("SOTPRIC2", "", True, ASCMAIN1.sql)
        If dst.Tables("SOTPRIC2").Rows.Count <> 0 Then
            For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                'DATETIME_STAMP = CDate(rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE"))
                rowSOTPRIC2.Item("ITEM_PRICE") = rowSOTPRIC2.Item("ITEM_NEW_PRICE")
                rowSOTPRIC2.Item("ITEM_NEW_PRICE") = DBNull.Value
                rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") = DBNull.Value
                Write_Audit_Trail(rowSOTPRIC2, "E")
            Next
            Update_Record_TDA("SOTPRIC2")
        End If

        ASCMAIN1.sql = "Select * from SOTPRIC2" & vbCrLf _
            & "  where ITEM_NEW_SRP_DATE <= '" & Format(CYPdt.AddDays(1), "dd-MMM-yyyy") & "'"
        Fill_Records("SOTPRIC2", "", True, ASCMAIN1.sql)
        If dst.Tables("SOTPRIC2").Rows.Count <> 0 Then
            For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                'DATETIME_STAMP = CDate(rowSOTPRIC2.Item("ITEM_NEW_SRP_DATE"))
                rowSOTPRIC2.Item("ITEM_SRP") = rowSOTPRIC2.Item("ITEM_NEW_SRP")
                rowSOTPRIC2.Item("ITEM_NEW_SRP") = DBNull.Value
                rowSOTPRIC2.Item("ITEM_NEW_SRP_DATE") = DBNull.Value
                Write_Audit_Trail(rowSOTPRIC2, "E")
            Next
            Update_Record_TDA("SOTPRIC2")
        End If



        ' Promo - Verified As Open flag

        ASCMAIN1.sql = "Update SPTCOOP1 X Set VERIFIED_AS_OPEN = NULL where VERIFIED_AS_OPEN = '1'"
        ASCDATA1.ExecuteSQL()


        ' Physical Inventory

        Dim SQLW As String = ""
        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"
        ASCMAIN1.sql = "" _
          & "Update ICTWHSE1 X Set WHSE_PHYS_STATUS = NULL" & ASCMAIN1.SQL_Add_WHERE(SQLW)
        ASCDATA1.ExecuteSQL()

        ' WHY DO WE NEED TO INITIALIZE THESE?
        'ASCDATA1.ExecuteSQL("Delete from ICTPHYC1")
        'ASCDATA1.ExecuteSQL("Delete from ICTPHYC2")


        TAC.SOCMAIN1.SetMonthlyCommissions()


        ' Set Inventory Cost Variances

        ' Reset TRAN_xV fields in ICTIREC2 (because of Re-Valuation) - ICTIREC2 contributions to xV amounts only

        ASCMAIN1.sql = $"Begin Declare Cursor C1 is
            Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO
            , Sum (ICTIREC2.QTY_REC * (ICTIREC2.PO_COST - ICTCOSTA.ITEM_COST_VCOST)) TRAN_PV
            , Sum (NVL(ICTIREC2.EXT_COST_MATLS,0) - ICTIREC2.QTY_REC * (
                    NVL(ICTCOSTA.ITEM_COST_MATLS,0) +
                    NVL(ICTCOSTA.ITEM_COST_LANDGI,0) +
                    NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) +
                    NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)
                    )) TRAN_MV
            , Sum (ICTIREC2.QTY_REC * (NVL(ICTIREC2.PO_COST_FRT,0) - NVL(ICTCOSTA.ITEM_COST_LANDG,0))) TRAN_FV
            , Sum (ICTIREC2.QTY_REC * (NVL(ICTIREC2.PO_COST_TRF,0) - NVL(ICTCOSTA.ITEM_COST_TOOLG,0))) TRAN_TV
                from ICTIREC2,ICTCOSTA
                where ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP
                  and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE
                  and ICTIREC2.OPS_YYYYPP = '{ASCMAIN1.CYP}'
            group by ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO;
            Begin For R1 in C1 Loop
            Update ICTIREC2 Set TRAN_PV = R1.TRAN_PV, TRAN_MV = R1.TRAN_MV, TRAN_FV = R1.TRAN_FV, TRAN_TV = R1.TRAN_TV
            where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;
            End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ' PO Receipts

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Update ICTIVAR1 Set PV_EXP = 0, MV_EXP = 0, CV_EXP = 0, FV_EXP = 0, TV_EXP = 0" & vbCrLf _
            & "  where OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ICTIREC2.ITEM_CODE, ICTIREC2.OPS_YYYYPP" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_PV) TRAN_PV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_MV) TRAN_MV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_CV) TRAN_CV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_FV) TRAN_FV" & vbCrLf _
            & ", Sum (ICTIREC2.TRAN_TV) TRAN_TV" & vbCrLf _
            & " from ICTIREC2,ICTCOSTA" & vbCrLf _
            & " where (NVL(TRAN_PV,0) <> 0 " & vbCrLf _
            & "    or NVL(TRAN_MV,0) <> 0 or NVL(TRAN_CV,0) <> 0 or NVL(TRAN_FV,0) <> 0 or NVL(TRAN_TV,0) <> 0)" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP" & vbCrLf _
            & "   and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & "   and ICTIREC2.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "group by ICTIREC2.ITEM_CODE, ICTIREC2.OPS_YYYYPP;" & vbCrLf _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set PV_EXP = R1.TRAN_PV, MV_EXP = R1.TRAN_MV, CV_EXP = R1.TRAN_CV" & vbCrLf _
            & "    , FV_EXP = R1.TRAN_FV, TV_EXP = R1.TRAN_TV" & vbCrLf _
            & "  where ITEM_CODE = R1.ITEM_CODE AND OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, PV_EXP, MV_EXP, CV_EXP, FV_EXP, TV_EXP)" & vbCrLf _
            & "   values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.TRAN_PV, R1.TRAN_MV, R1.TRAN_CV, R1.TRAN_FV, R1.TRAN_TV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()

        ' PV from AP Invoicing

        Dim sqlAPTINVH5 As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP" & vbCrLf _
            & ", Sum (APTINVH5.VAR_AMT) PV" & vbCrLf _
            & " from APTINVH5,ICTIREC2,APTINVH1" & vbCrLf _
            & " where ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO = APTINVH5.RECEIPT_LNO" & vbCrLf _
            & "   and APTINVH1.REGISTER_IND = '1'" & vbCrLf _
            & "   and APTINVH5.VAR_AMT <> 0" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP;" & vbCrLf

        ASCMAIN1.sql = sqlAPTINVH5 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Update ICTIVAR1 " & vbCrLf _
            & "Set PV_EXP = NVL(PV_EXP,0) + R1.PV" & vbCrLf _
            & " where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & "If SQL%NOTFOUND THEN" & vbCrLf _
            & "Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, PV_EXP)" & vbCrLf _
            & " Values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.PV);" & vbCrLf _
            & "End If;" & vbCrLf _
            & "End Loop; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()

        sqlAPTINVH5 = Replace(sqlAPTINVH5, "ICTIREC2.ITEM_CODE, APTINVH1.OPS_YYYYPP", "ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTINVH5 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & "Update ICTIREC2 " & vbCrLf _
            & "Set TRAN_PV = NVL(TRAN_PV,0) + NVL(R1.PV,0)" & vbCrLf _
            & " where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;" & vbCrLf
        ASCDATA1.ExecuteSQL()


        ' FV & TV from Invoicing

        Dim sqlAPTACRC1 As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ITEM_CODE, OPS_YYYYPP" & vbCrLf _
            & ", Sum (CASE WHEN ACCRUAL_CODE = 'FRT' THEN VAR ELSE 0 END) FV" & vbCrLf _
            & ", Sum (CASE WHEN ACCRUAL_CODE = 'TRF' THEN VAR ELSE 0 END) TV" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select APTACRC1.ITEM_CODE, APTINVH1.OPS_YYYYPP, APTACRC1.ACCRUAL_CODE" & vbCrLf _
            & " , APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO" & vbCrLf _
            & ", Sum (NVL(APTINVH7.TOTAL_INV,0) - NVL(APTINVH7.TOTAL_ACC,0)) VAR" & vbCrLf _
            & " from APTINVH7,APTACRC1,APTINVH1" & vbCrLf _
            & " where APTACRC1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
            & "   and APTINVH1.REGISTER_IND = '1'" & vbCrLf _
            & "   and (APTACRC1.ACCRUAL_CODE = 'FRT' or APTACRC1.ACCRUAL_CODE = 'TRF')" & vbCrLf _
            & "   and NVL(APTACRC1.PPD_IND,'0') = '0' and PPD_MATCHED_XNO is Null" & vbCrLf _
            & "   and APTACRC1.ITEM_CODE is Not Null" & vbCrLf _
            & "   and NVL(APTINVH7.TOTAL_INV,0) - NVL(APTINVH7.TOTAL_ACC,0) <> 0" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO = APTINVH7.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by APTACRC1.ITEM_CODE, APTINVH1.OPS_YYYYPP, APTACRC1.ACCRUAL_CODE" & vbCrLf _
            & ", APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO)" & vbCrLf _
            & " group by ITEM_CODE, OPS_YYYYPP;" & vbCrLf

        ASCMAIN1.sql = sqlAPTACRC1 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set FV_EXP = NVL(FV_EXP,0) + R1.FV, TV_EXP = NVL(TV_EXP,0) + R1.TV" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, FV_EXP, TV_EXP)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.OPS_YYYYPP, R1.FV, R1.TV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        sqlAPTACRC1 = Replace(sqlAPTACRC1, "ITEM_CODE, OPS_YYYYPP", "RECEIPT_NO, RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTACRC1 _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIREC2 " & vbCrLf _
            & "  Set TRAN_FV = NVL(TRAN_FV,0) + R1.FV, TRAN_TV = NVL(TRAN_TV,0) + NVL(R1.TV,0)" & vbCrLf _
            & "   where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ' TV from PPD Match

        Dim sqlAPTACRCM As String = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select ITEM_CODE, OPS_YYYYPP_MATCHED, Sum (COST_VAR_ITEM) TV" & vbCrLf _
            & " from APTACRC1" & vbCrLf _
            & " where OPS_YYYYPP_MATCHED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and NVL(PPD_IND,'0') = '0'" & vbCrLf _
            & " group by ITEM_CODE, OPS_YYYYPP_MATCHED HAVING Sum (COST_VAR_ITEM) <> 0;" & vbCrLf

        ASCMAIN1.sql = sqlAPTACRCM _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIVAR1 " & vbCrLf _
            & "  Set TV_EXP = NVL(TV_EXP,0) + NVL(R1.TV,0)" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = R1.OPS_YYYYPP_MATCHED;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "  Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, TV_EXP)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.OPS_YYYYPP_MATCHED, R1.TV);" & vbCrLf _
            & " End If;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        sqlAPTACRCM = Replace(sqlAPTACRCM, "ITEM_CODE, OPS_YYYYPP_MATCHED", "RECEIPT_NO, RECEIPT_LNO")
        ASCMAIN1.sql = sqlAPTACRCM _
            & "Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTIREC2 " & vbCrLf _
            & "  Set TRAN_TV = NVL(TRAN_TV,0) + NVL(R1.TV,0)" & vbCrLf _
            & "   where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ' Reconciling Variances

        'Select Case COST_CATGY_CODE, sum (PV_EXP), sum (MV_EXP),  sum (FV_EXP), sum (TV_EXP)
        'From ictivar1, ictitem1
        'Where ops_yyyypp = '202510'
        'And ictitem1.item_code = ictivar1.item_code
        'Group By COST_CATGY_CODE;

        'Select Case COST_CATGY_CODE, sum (tran_Pv), sum (tran_Mv), sum (tran_Fv), sum (tran_tv)
        'From ictirec2
        'Where ops_yyyypp = '202510'
        'group by COST_CATGY_CODE;


        ' The difference between the TRAN_xV fields And the xV_EXP fields Is ok because a
        ' variance recorded in AP in ASCMAIN1.CYP might impact a receipt from a prior month


        ' Copy Future % to Current % of MSRP for Retail Cost Classes

        'ASCMAIN1.sql = "Update ICTCCLS1 Set COST_BASE_PCT_OF_MSRP = COST_BASE_PCT_OF_MSRP_FUT where COST_BASIS = 'R' and NVL(COST_BASE_PCT_OF_MSRP,0) <> NVL(COST_BASE_PCT_OF_MSRP_FUT,0)"
        'ASCDATA1.ExecuteSQL()

        ' Archive Cost Lists - is there a time when we would stop doing this?  like maybe when a Cost List is retired or inactivated?

        'ASCMAIN1.sql = $"Insert into ICTCLST4 Select '{ASCMAIN1.CYP}' OPS_YYYYPP, ICTCLST2.* from ICTCLST2"
        'ASCDATA1.ExecuteSQL()

        ' Copy Season %'s by Week

        If ASCMAIN1.CYP = "06" Or ASCMAIN1.CYP = "12" Then
            Dim SEASON_TYPE As String = "S"
            If ASCMAIN1.CYP = "12" Then SEASON_TYPE = "F"
            ASCMAIN1.sql = "Select '" & Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) + 1, "0000") & SEASON_TYPE & "', WEEK_NO, PCT_M, PCT_W" _
                & " from SPTMXWS0 " _
                & " where SEASON_CODE = '" & Mid(ASCMAIN1.CYP, 1, 4) & SEASON_TYPE & "'"
            ASCMAIN1.sql = "Insert into SPTMXWS0 " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, True)
        End If

        If ASCMAIN1.CLIENT = "INT" Then

            ' Demo Commissions
            ASCMAIN1.Progress("Demo Commissions Snapshots", "")

            ASCMAIN1.sql = "Insert into SPTDCOMH Select '" & RYP & "' OPS_YYYYPP_H, SPTDCOMB.* from SPTDCOMB WHERE NVL(SPTDCOMB.AMT_COMM,0) - NVL(SPTDCOMB.AMT_COMM_PAID,0) <> 0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into SPTDCOMI Select '" & RYP & "' OPS_YYYYPP_H, SPTDCOMC.* from SPTDCOMC"
            ASCDATA1.ExecuteSQL()
        End If

        If ASCMAIN1.CLIENT = "INT" Then
            ' DO NOTHING 
        Else
            ASCMAIN1.sql = "Insert into SOTRMAF3 Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP" & vbCrLf _
                & ", SOTRMAF2.* from SOTRMAF2,SOTRMAF1" & vbCrLf _
                & " where SOTRMAF1.RA_NO = SOTRMAF2.RA_NO" & vbCrLf _
                & "   and SOTRMAF1.RA_START_DATE <= '" & Format(CYPdt, "dd-MMM-yyyy") & "'" & vbCrLf _
                & "   and SOTRMAF2.RA_QTY_OPEN <> 0"
            ASCDATA1.ExecuteSQL()
        End If

        ' GL Control Account Subsidiary Snapshots
        Load_GLREC()

        ' Close Period
        ASCMAIN1.Progress("Updating Period Control Record", "")
        ASCMAIN1.sql = "Update ASTPCTL1 set CURR_YEAR = '" & Mid$(NYP, 1, 4) & "'," & vbCrLf _
            & " CURR_PERIOD = '" & Mid$(NYP, 5, 2) & "'," & vbCrLf _
            & " PRD_CLOSE_IND = Null"
        ASCDATA1.ExecuteSQL()

        ' Purge SQL Transactions
        ASCMAIN1.Progress("Cleaning up SQL Transaction History", "")
        ASCMAIN1.sql = "Delete from ASTSQLX1"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Load_GLREC()

        Dim cols As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        ASCMAIN1.sql = "Delete from GLTCREC1"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from GLTCREC2"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from GLTCREC4"
        ASCDATA1.ExecuteSQL()

        ' Inventory On Hand
        ASCMAIN1.Progress("Subsidiary Snapshots", "Inventory On Hand")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'IC', 'Inventory On Hand', 'D' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'IC', PROD_CODE, PROD_DESC from ICTPROD1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'IC' CREC_TYPE_CODE, ICTITEM1.PROD_CODE PROD_CODE" & vbCrLf _
            & ", ICTCOST1.ACCT_CODE_ONH ACCT_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", NULL DETL_CTL_TYPE, NULL DETL_CTL_NO, 'V' DETL_CVX_TYPE, ICTITEM1.COST_CATGY_CODE DETL_CVX_NO" & vbCrLf _
            & ", SUM(NVL(ICTITEM1.ITEM_COST_STD,0) * NVL(ICTSTAT5.WHSE_QTY_ON_HAND,0)) CREC_AMT" & vbCrLf _
            & " from ICTSTAT5,ICTITEM1,ICTCOST1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = ICTSTAT5.ITEM_CODE" & vbCrLf _
            & "   and ICTSTAT5.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & " group by ICTITEM1.PROD_CODE, ICTCOST1.ACCT_CODE_ONH" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE" & vbCrLf
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ' Accrued Purchases
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Purchases")
        ASCMAIN1.sql = "INSERT INTO GLTCREC1 SELECT 'ICP', 'Accrued Purchases', 'C' FROM DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ICP', PROD_CODE, PROD_DESC from ICTPROD1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICP' CREC_TYPE_CODE," & vbCrLf _
            & "       ICTITEM1.COST_CATGY_CODE CREC_CLASS_CODE, '" & ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_PURCH") & "' ACCT_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE," & vbCrLf _
            & "       'R' DETL_CTL_TYPE, ICTIREC1.RECEIPT_NO DETL_CTL_NO, 'V' DETL_CVX_TYPE, ICTIREC1.VEND_CODE DETL_CVX_NO," & vbCrLf _
            & "       SUM(ICTIREC2.PO_COST * (NVL(ICTIREC2.QTY_REC,0) - NVL(ICTIREC2.QTY_INV,0))) CREC_AMT" & vbCrLf _
            & " from ICTIREC1,ICTIREC2,ICTITEM1" & vbCrLf _
            & "  where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & "   and ICTIREC2.ACCRUAL_STATUS = '0'" & vbCrLf _
            & " group by ICTITEM1.COST_CATGY_CODE, ICTIREC1.RECEIPT_NO, ICTIREC1.VEND_CODE"
        ASCMAIN1.sql = "INSERT INTO GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()


        ' Accrued Freight
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Freight")
        ASCMAIN1.sql = "INSERT INTO GLTCREC1 SELECT 'ICF', 'Accrued Freight', 'C' FROM DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ICF', VEND_CODE, VEND_NAME from APTVEND1 where VEND_CODE in (Select VEND_CODE_ACC from APTACRC1 where CTL_STATUS = '0')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICF' CREC_TYPE_CODE," & vbCrLf _
            & "       APTACRC1.COST_CATGY_CODE CREC_CLASS_CODE, '" & ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_LANDG") & "' ACCT_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE," & vbCrLf _
            & "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE," & vbCrLf _
            & "       'R' DETL_CTL_TYPE, APTACRC1.RECEIPT_NO DETL_CTL_NO, 'V' DETL_CVX_TYPE, APTACRC1.VEND_CODE_ACC DETL_CVX_NO," & vbCrLf _
            & "       SUM(NVL(APTACRC1.COST_ACC,0) - NVL(APTACRC1.COST_ACT,0)) CREC_AMT" & vbCrLf _
            & " from APTACRC1" & vbCrLf _
            & "  where APTACRC1.CTL_STATUS = '0' AND APTACRC1.ACCRUAL_CODE = 'FRT'" & vbCrLf _
            & " group by APTACRC1.COST_CATGY_CODE, APTACRC1.RECEIPT_NO, APTACRC1.VEND_CODE_ACC"
        ASCMAIN1.sql = "INSERT INTO GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        Dim cols5 As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CTL_LNO,CREC_QTY,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICF' CREC_TYPE_CODE," & vbCrLf _
            & "       ICTITEM1.COST_CATGY_CODE CREC_CLASS_CODE,'R' DETL_CTL_TYPE, " & vbCrLf _
            & "       ICTIREC2.RECEIPT_NO DETL_CTL_NO, ICTIREC2.RECEIPT_LNO DETL_CTL_LNO, " & vbCrLf _
            & "       NVL(ICTIREC2.QTY_REC,0) CREC_QTY," & vbCrLf _
            & "       'C' DETL_CVX_TYPE, APTACRC1.CTL_NO DETL_CVX_NO," & vbCrLf _
            & "       APTACRC1.COST_ACC CREC_AMT" & vbCrLf _
            & " from APTACRC1,ICTIREC2,ICTITEM1" & vbCrLf _
            & "  where APTACRC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "    and APTACRC1.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & "    and ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
            & "    and APTACRC1.CTL_STATUS = '0' and APTACRC1.ACCRUAL_CODE = 'FRT'"
        ASCMAIN1.sql = "INSERT INTO GLTCREC5 " & cols5 & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ' Accrued Tariff
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Tariff")
        ASCMAIN1.sql = "INSERT INTO GLTCREC1 SELECT 'ICT', 'Accrued Tariff', 'C' FROM DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ICT', VEND_CODE, VEND_NAME from APTVEND1 where VEND_CODE in (Select VEND_CODE_ACC from APTACRC1 where ACCRUAL_CODE = 'TRF' AND (CTL_STATUS = '0' OR (CTL_STATUS = '1' AND PPD_IND = '1')))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICT' CREC_TYPE_CODE," & vbCrLf &
               "       APTACRC1.COST_CATGY_CODE CREC_CLASS_CODE, '" & ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_TOOLG") & "' ACCT_CODE," & vbCrLf &
               "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE," & vbCrLf &
               "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE," & vbCrLf &
               "       '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE," & vbCrLf &
               "       'R' DETL_CTL_TYPE, APTACRC1.RECEIPT_NO DETL_CTL_NO, 'V' DETL_CVX_TYPE, APTACRC1.VEND_CODE_ACC DETL_CVX_NO," & vbCrLf &
               "       SUM(NVL(APTACRC1.COST_ACC,0) - NVL(APTACRC1.COST_ACT,0)) CREC_AMT" & vbCrLf &
               " from APTACRC1" & vbCrLf &
               " where APTACRC1.ACCRUAL_CODE = 'TRF'" & vbCrLf &
               "   AND (CTL_STATUS = '0' or (CTL_STATUS = '1' AND NVL(PPD_IND,'0') = '1' AND NVL(PPD_MATCHED,'0') = '0'))" & vbCrLf &
               " group by APTACRC1.COST_CATGY_CODE, APTACRC1.RECEIPT_NO, APTACRC1.VEND_CODE_ACC"
        ASCMAIN1.sql = "INSERT INTO GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        Dim cols6 As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CTL_LNO,CREC_QTY,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        ASCMAIN1.sql = "SELECT '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ICT' CREC_TYPE_CODE," & vbCrLf &
               "       APTACRC1.COST_CATGY_CODE CREC_CLASS_CODE, 'R' DETL_CTL_TYPE," & vbCrLf &
               "       APTACRC1.RECEIPT_NO DETL_CTL_NO, APTACRC1.RECEIPT_LNO DETL_CTL_LNO," & vbCrLf &
               "       NVL(ICTIREC2.QTY_REC,0) CREC_QTY," & vbCrLf &
               "       'C' DETL_CVX_TYPE, APTACRC1.CTL_NO DETL_CVX_NO," & vbCrLf &
               "       NVL(APTACRC1.COST_ACC,0) - NVL(APTACRC1.COST_ACT,0) CREC_AMT" & vbCrLf &
               " from APTACRC1,ICTIREC2" & vbCrLf &
               " where APTACRC1.ACCRUAL_CODE = 'TRF'" & vbCrLf &
               "   and (APTACRC1.CTL_STATUS = '0' or (APTACRC1.CTL_STATUS = '1' AND NVL(APTACRC1.PPD_IND,'0') = '1' AND NVL(APTACRC1.PPD_MATCHED,'0') = '0'))" & vbCrLf &
               "   and ICTIREC2.RECEIPT_NO (+) = APTACRC1.RECEIPT_NO" & vbCrLf &
               "   and ICTIREC2.RECEIPT_LNO (+) = APTACRC1.RECEIPT_LNO"

        ASCMAIN1.sql = "INSERT INTO GLTCREC5 " & cols6 & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()


        ' Accounts Receivable
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accounts Receivable")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'AR', 'Accounts Receivable', 'D' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'AR', POST_CODE, POST_DESC from ARTPOST1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'AR' CREC_TYPE_CODE, ARTOPEN1.POST_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", ARTPOST1.ACCT_CODE ACCT_CODE" & vbCrLf _
            & ", ARTOPEN1.SEG2_CODE, ARTOPEN1.SEG3_CODE, ARTOPEN1.SEG4_CODE" & vbCrLf _
            & ", ARTOPEN1.INV_TYPE DETL_CTL_TYPE, ARTOPEN1.INV_NUM DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, ARTOPEN1.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE CREC_AMT" & vbCrLf _
            & " from ARTOPEN1,ARTPOST1" & vbCrLf _
            & " where ARTPOST1.POST_CODE (+) = ARTOPEN1.POST_CODE" & vbCrLf _
            & "   and NVL(ARTOPEN1.OPS_YYYYPP,'" & ASCMAIN1.CYP & "') <= '" & ASCMAIN1.CYP & "'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTCREC4 Select 'C', CUST_CODE, CUST_NAME from ARTCUST1"
        ASCDATA1.ExecuteSQL()

        ' Accounts Payable
        ASCMAIN1.Progress("Subsidiary Snapshots", "Accounts Payable")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'AP', 'Accounts Payable', 'C' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'AP', POST_CODE, POST_DESC from APTPOST1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'AP' CREC_TYPE_CODE, APTINVH1.POST_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", APTPOST1.ACCT_CODE ACCT_CODE" & vbCrLf _
            & ", APTINVH1.SEG2_CODE, APTINVH1.SEG3_CODE, APTINVH1.SEG4_CODE" & vbCrLf _
            & ", APTINVH1.INV_TYPE DETL_CTL_TYPE, APTINVH1.VOUCHER_NO DETL_CTL_NO" & vbCrLf _
            & ", 'V' DETL_CVX_TYPE, APTINVH1.VEND_CODE DETL_CVX_NO" & vbCrLf _
            & ", APTINVH1.INV_BALANCE CREC_AMT" & vbCrLf _
            & " from APTINVH1,APTPOST1" & vbCrLf _
            & " where APTPOST1.POST_CODE (+) = APTINVH1.POST_CODE" & vbCrLf _
            & "   and (APTINVH1.INV_STATUS = 'O' or APTINVH1.INV_STATUS = 'H')"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTCREC4 Select 'V', VEND_CODE, VEND_NAME from APTVEND1"
        ASCDATA1.ExecuteSQL()




        ' Promo Expenses Commissions
        ASCMAIN1.Progress("Subsidiary Snapshots", "Promo Expenses")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'PX', 'Promo Expenses', 'C' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'PX', EXPENSE_TYPE_CODE, EXPENSE_TYPE_DESC from SPTTYPE1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'PX' CREC_TYPE_CODE, SPTCOOP1.EXPENSE_TYPE_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", SPTPARM1.SP_PARM_PROMO_ACCT_CODE_ACC ACCT_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", 'C' DETL_CTL_TYPE, SPTCOOP1.AUTH_NO DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, SPTCOOP1.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", CASE WHEN (NVL(QTY,0) = 0 OR NVL(VEHICLE_CPM,0) = 0) AND NVL(OTHER_COST,0) = 0 THEN 0 ELSE ROUND(NVL(SPTCOOP1.OPEN_AMT,0) * NVL(SPTCOOP3.DIST_AMT,0) / (NVL(QTY,0) * NVL(VEHICLE_CPM,0) / 1000 + NVL(OTHER_COST,0)),2) END CREC_AMT" & vbCrLf _
            & " from SPTCOOP1,SPTCOOP3,SPTPARM1" & vbCrLf _
            & " where SPTCOOP1.STATUS_CODE = 'O'" & vbCrLf _
            & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO" & vbCrLf _
            & "   and NVL(SPTCOOP1.OPS_YYYYPP_ACCRUE,SPTCOOP1.OPS_YYYYPP) <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and SPTPARM1.SP_PARM_KEY = 'Z'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Insert into GLTCREC4 Select 'C', CUST_CODE, CUST_NAME from ARTCUST1"
        'ASCDATA1.ExecuteSQL()




        ' Demo Commissions
        ASCMAIN1.Progress("Subsidiary Snapshots", "Demo Commissions")
        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'DC', 'Demo Commissions', 'C' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'DC', BRAND_CODE, BRAND_NAME from ICTBRAN1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'DC' CREC_TYPE_CODE, SPTDCOMC.BRAND_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", SPTPARM1.SP_PARM_DEMO_ACCT_CODE_ACC ACCT_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", 'D' DETL_CTL_TYPE, SPTDCOMC.ACC_CTL_NO DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, SPTDCOMC.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", NVL(SPTDCOMC.AMT_COMM,0)-NVL(SPTDCOMC.AMT_COMM_OFFSET,0) CREC_AMT" & vbCrLf _
            & " from SPTDCOMC,SPTPARM1" & vbCrLf _
            & " where SPTDCOMC.PYMT_NO is Null" & vbCrLf _
            & "   and SPTPARM1.SP_PARM_KEY = 'Z'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Insert into GLTCREC4 Select 'C', CUST_CODE, CUST_NAME from ARTCUST1"
        'ASCDATA1.ExecuteSQL()


        If ASCMAIN1.CLIENT = "AHA" Then
            ' LEAVING THIS IN HERE BECAUSE i KEEP FORGETTING THAT THIS NEEDS TO BE DONE AFTER THE ACCRUAL JOURNAL IS RUN
            ' - WELL NOT EXACTLY - LEAVING THIS HERE AND UNREMMING IT BECAUSE THIS NEEDS TO RECORD WHAT IS OPEN AT MONTH END,
            '   AND THEN WE ADD NEW ACCRUALS WHEN THE ACCRUAL JOURNAL IS RUN
            '    ' Advertising Commissions
            ASCMAIN1.Progress("Subsidiary Snapshots", "Advertising & Commissions")
            ASCMAIN1.sql = "Insert into GLTCREC1 Select 'AC', 'Advertising & Commissions', 'C' from DUAL"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into GLTCREC2 Select 'AC', ASP_CODE, ASP_DESC from SPTACOM0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'AC' CREC_TYPE_CODE, SPTACOMC.ASP_CODE CREC_CLASS_CODE" & vbCrLf _
                & ", NVL(SPTACOM0.ASP_ACCT_CODE_ACC,SPTPARM1.SP_PARM_ASP_ACCT_CODE_ACC) ACCT_CODE" & vbCrLf _
                & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", 'D' DETL_CTL_TYPE, SPTACOMC.ACC_CTL_NO DETL_CTL_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SPTACOMC.CUST_CODE DETL_CVX_NO" & vbCrLf _
                & ", NVL(SPTACOMC.AMT_COMM,0)-NVL(SPTACOMC.AMT_COMM_OFFSET,0) CREC_AMT" & vbCrLf _
                & " from SPTACOMC,SPTACOM0,SPTPARM1" & vbCrLf _
                & " where SPTACOMC.PYMT_NO is Null" & vbCrLf _
                & "   and SPTACOM0.ASP_CODE = SPTACOMC.ASP_CODE" & vbCrLf _
                & "   and SPTPARM1.SP_PARM_KEY = 'Z'"
            ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        End If

        If ASCMAIN1.CLIENT = "AHA" Then

            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
            Dim DETL_CTL_DATE_LAST As Date = rowGLTPARM2.Item("PRD_END_DATE")

            ' Accrued Returns Authorizations - Revenue

            ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Returns")
            ASCMAIN1.sql = "Insert into GLTCREC1 SELECT 'ARR1', 'Accrued Returns - Revenue', 'C' FROM DUAL"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ARR1', PROD_CODE, PROD_DESC from ICTPROD1"
            ASCDATA1.ExecuteSQL()

            '                & ", Sum (-1 * NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) CREC_AMT" & vbCrLf _

            ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ARR1' CREC_TYPE_CODE" & vbCrLf _
                & ", ICTITEM1.PROD_CODE CREC_CLASS_CODE" & vbCrLf _
                & ", SOTPARM1.SO_PARM_RA_REC_ACCT_CODE ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", 'Z' DETL_CTL_TYPE, SOTRMAF1.RA_NO DETL_CTL_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTRMAF1.CUST_CODE DETL_CVX_NO" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) CREC_AMT" & vbCrLf _
                & " from SOTRMAF2,ICTITEM1,SOTRMAF1,SOTPARM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE " & vbCrLf _
                & "   and SOTRMAF1.RA_NO = SOTRMAF2.RA_NO" & vbCrLf _
                & "   and NVL(SOTRMAF2.RA_QTY_OPEN,0) <> 0" & vbCrLf _
                & "   and SOTPARM1.SO_PARM_KEY = 'Z'" & vbCrLf _
                & "   and SOTRMAF1.RA_START_DATE <= '" & Format(DETL_CTL_DATE_LAST, "dd-MMM-yyyy") & "'" & vbCrLf _
                & " group by ICTITEM1.PROD_CODE, SOTRMAF1.RA_NO, SOTRMAF1.CUST_CODE" & vbCrLf _
                & ", SOTPARM1.SO_PARM_RA_REC_ACCT_CODE"
            ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()

            ' Accrued Returns Authorizations - Invty

            ASCMAIN1.Progress("Subsidiary Snapshots", "Accrued Returns")
            ASCMAIN1.sql = "Insert into GLTCREC1 SELECT 'ARR2', 'Accrued Returns - CGS', 'D' FROM DUAL"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into GLTCREC2 Select 'ARR2', PROD_CODE, PROD_DESC from ICTPROD1"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'ARR2' CREC_TYPE_CODE" & vbCrLf _
                & ", ICTITEM1.PROD_CODE CREC_CLASS_CODE" & vbCrLf _
                & ", SOTPARM1.SO_PARM_RA_INV_ACCT_CODE ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", 'Z' DETL_CTL_TYPE, SOTRMAF1.RA_NO DETL_CTL_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTRMAF1.CUST_CODE DETL_CVX_NO" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(ICTITEM1.ITEM_COST_STD,0)) CREC_AMT" & vbCrLf _
                & " from SOTRMAF2,ICTITEM1,SOTRMAF1,SOTPARM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE " & vbCrLf _
                & "   and SOTRMAF1.RA_NO = SOTRMAF2.RA_NO" & vbCrLf _
                & "   and NVL(SOTRMAF2.RA_QTY_OPEN,0) <> 0" & vbCrLf _
                & "   and SOTPARM1.SO_PARM_KEY = 'Z'" & vbCrLf _
                & "   and SOTRMAF1.RA_START_DATE <= '" & Format(DETL_CTL_DATE_LAST, "dd-MMM-yyyy") & "'" & vbCrLf _
                & " group by ICTITEM1.PROD_CODE, SOTRMAF1.RA_NO, SOTRMAF1.CUST_CODE" & vbCrLf _
                & ", SOTPARM1.SO_PARM_RA_INV_ACCT_CODE"
            ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()

            ' Deferred Revenue Recognition
            'ASCMAIN1.Progress("Subsidiary Snapshots", "Deferred Revenue Recognition")
            'ASCMAIN1.sql = "Insert into GLTCREC1 Select 'OPDR', 'Deferred Revenue Recognition', 'D' from DUAL"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Insert into GLTCREC2 Select 'OPDR', 'SLS', 'Sales' from DUAL"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Insert into GLTCREC2 Select 'OPDR', 'CGS', 'Cost of Goods Sold' from DUAL"
            'ASCDATA1.ExecuteSQL()

            'Dim sqlw As String = ""
            'sqlw &= " AND SOTORDR1.FRT_TERMS = 'DEL'" & vbCrLf
            'sqlw &= " AND SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP0 & "'" & vbCrLf
            'sqlw &= " AND (SOTINVH1.ORDR_YYYYPP_DEL IS NULL" & vbCrLf
            'sqlw &= " OR SOTINVH1.ORDR_YYYYPP_DEL > '" & RYP0 & "')" & vbCrLf

            'ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'OPDR' CREC_TYPE_CODE, 'SLS' CREC_CLASS_CODE" & vbCrLf _
            '    & ", SOTPARM1.SO_PARM_ACCT_CODE_DEF_SLS ACCT_CODE" & vbCrLf _
            '    & ", '000' SEG2_CODE, '000' SEG3_CODE, '000' SEG4_CODE" & vbCrLf _
            '    & ", SOTINVH1.INV_TYPE DETL_CTL_TYPE, SOTINVH1.INV_NO DETL_CTL_NO" & vbCrLf _
            '    & ", 'C' DETL_CVX_TYPE, SOTINVH1.CUST_CODE DETL_CVX_NO" & vbCrLf _
            '    & ", SOTINVH1.INV_SALES CREC_AMT" & vbCrLf _
            '    & " from SOTINVH1, SOTPARM1, SOTORDR1" & vbCrLf _
            '    & " where SOTPARM1.SO_PARM_KEY = 'Z' AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO " & sqlw
            'ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, 'OPDR' CREC_TYPE_CODE, 'CGS' CREC_CLASS_CODE" & vbCrLf _
            '    & ", SOTPARM1.SO_PARM_ACCT_CODE_DEF_CGS ACCT_CODE" & vbCrLf _
            '    & ", '000' SEG2_CODE, '000' SEG3_CODE, '000' SEG4_CODE" & vbCrLf _
            '    & ", SOTINVH1.INV_TYPE DETL_CTL_TYPE, SOTINVH1.INV_NO DETL_CTL_NO" & vbCrLf _
            '    & ", 'C' DETL_CVX_TYPE, SOTINVH1.CUST_CODE DETL_CVX_NO" & vbCrLf _
            '    & ", SOTINVH1.INV_COGS CREC_AMT" & vbCrLf _
            '    & " from SOTINVH1, SOTPARM1, SOTORDR1" & vbCrLf _
            '    & " where SOTPARM1.SO_PARM_KEY = 'Z' AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO" & sqlw
            'ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Insert into GLTCREC4 Select 'C', CUST_CODE, CUST_NAME from ARTCUST1"
            'ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Purge_Files()
        Get_PARM("ASTPARM1")

        Dim AS_PARM_ARCHIVE_FOLDER As String = ROWs("ASTPARM1").Item("AS_PARM_ARCHIVE_FOLDER") & "\"
        Dim AS_PARM_REPORTS_ARCHIVE_DAYS As Integer = Val(ROWs("ASTPARM1").Item("AS_PARM_REPORTS_ARCHIVE_DAYS") & "")

        ' Do not Purge Reports if Days = -1
        If AS_PARM_REPORTS_ARCHIVE_DAYS = -1 Then Exit Sub

        Dim PURGE_DATE As String = Format$(Now.AddDays(-1 * AS_PARM_REPORTS_ARCHIVE_DAYS), "dd-MMM-yyyy")

        Dim sql As String = "Select * from ASTSPRF1 where REPORT_DATE <= '" & PURGE_DATE & "'"
        For Each rowASTSPRF1 As DataRow In ASCDATA1.GetDataTable(sql, "ASTSPRF1").Rows
            On Error Resume Next
            Dim FILENAME As String = AS_PARM_ARCHIVE_FOLDER _
                                   & "Reports\" _
                                   & rowASTSPRF1.Item("REPORT_NO") _
                                   & "." & ROWs("ASTPARM1").Item("AS_PARM_REPORTS_SFX")

            ' NEW NAMING CONVENTION TO STORE FEWER REPORTS IN EACH FOLDER
            
            FILENAME = ASCMAIN1.DBS_COMPANY & "_" & rowASTSPRF1.Item("REPORT_NO") & ".RPT"
            FILENAME = ASCMAIN1.Folders("Archive") & "Reports\" & Mid(FILENAME, 1, 3) & "\" & Mid(FILENAME, 5, 5) & "\" & FILENAME

            My.Computer.FileSystem.DeleteFile(FILENAME)
            On Error GoTo 0
            rowASTSPRF1.Delete()
        Next
    End Sub

End Class