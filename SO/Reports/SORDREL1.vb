Public Class SORDREL1
    'Dim sql_pick As String          ' where clause for all SOTPICK1 in scope of the De-Release with PICK_STATUS of P or C
    'Dim sql_pick_all As String      ' where clause for all SOTPICK1 in scope of the De-Release

    Dim SOTPICK1 As String
    Dim SOTSHIP1 As String
    Dim dederelease_PTs As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            chkDeDeRelease.Visible = True
        End If
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        Dim sql_pick As String = Set_filter_for_Pick_Tickets_to_De_Release()

        If dederelease_PTs Then
            Stop
            If ASCMAIN1.Running_in_VS Then
                Stop
                ASCMAIN1.sql = "Select * from SOTPICK1_DDR" ' NEED TO SET UP THIS SQL STATEMENT TO PREPARE THE PTS TO BE DE-RELEASED
            Else
                Stop
            End If

        Else
            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & ASCMAIN1.SQL_Add_WHERE(sql_pick)
        End If

        SOTPICK1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1 & " Add Primary Key (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTPICK1 & "_1 on " & SOTPICK1 & " (SHIP_BOL_NO)")

        ASCMAIN1.sql = "Select * from " & SOTPICK1
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", 1))

        ASCMAIN1.sql = "Select SOTSHIP1.*,SOTORDR0.CUST_CODE from SOTSHIP1,SOTORDR0" _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
            & " and SOTSHIP1.SHIP_BOL_NO in " _
            & " (Select DISTINCT SHIP_BOL_NO from " & SOTPICK1 & ")"
        SOTSHIP1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP1 & " Add Primary Key (SHIP_BOL_NO)")

        '& " (Select DISTINCT SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")"

        ASCMAIN1.sql = "Select * from " & SOTSHIP1
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIP1", 1))

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        dst.Tables("SOTSHIP1").Columns.Add("PICK_CNT", GetType(System.Int32), "COUNT(CHILD.PICK_NO)")

        ASCMAIN1.sql = "Select Count (*) from " & SOTPICK1
        Dim C As Int64 = Val(ASCDATA1.GetDataValue)
        If C > 100 Then
            If MsgBox("An Excessive number of Pick Tickets are queued up to be De-Released" _
                      & vbCrLf & vbCrLf & "Are you sure that you want to Continue with the De-Release process?", _
                      MsgBoxStyle.YesNo, "Verfication - Over 100 Pick Tickets will be De-Released") = MsgBoxResult.No Then
                RWU = "N"
                MsgBox("Report will print, but Update will be Disabled", MsgBoxStyle.OkOnly, "Verification")
            End If
        End If

    End Sub

    Public Overrides Sub Print_Report()
        If dederelease_PTs Then
            RPT_TITLE = "Sales Order De-De-Release"
            SUBT = "(Special Update)"
        End If
        Generate_Report(RPT, RPT_TITLE, SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Proceed"
                Dim C As Integer = tblASTDSQLA.Select("ISNULL(CODE_VALUES,'') <> ''").Length ' tblASTDSQLA.Select("CODE_VALUES IS NOT NULL").Length


                If C < 1 Then
                    EMsg &= vbCr & "You Must Specify at least 1 Pick Batch, Ship BOL No, or Pick Ticket"
                ElseIf C > 1 Then
                    EMsg &= vbCr & "You Cannot Mix Pick Batches, Ship BOL No's and Pick Tickets in a Single Execution"
                End If

                If tblASTDSQLA.Select("EXCLUDE = '1'").Length <> 0 Then
                    EMsg &= vbCr & "You may not use Exclusion on any Filter for De-Release"
                End If

                Dim sql_pick As String = Set_filter_for_Pick_Tickets_to_De_Release()

                If SQLA("PICK_NO") <> "" Then
                    ASCMAIN1.sql = "Select Count (*) from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick)
                    If Val(ASCDATA1.GetDataValue) = 0 Then
                        EMsg &= vbCr & "No Pick Tickets to De-Release"
                    End If
                    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" _
                        & " and SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")"
                    If Val(ASCDATA1.GetDataValue) <> 0 Then
                        EMsg &= vbCr & "Some Shipments Selected are Not In Pick"
                    End If
                End If

                If SQLA("SHIP_BOL_NO") <> "" Then
                    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" & SQL_in("SHIP_BOL_NO")
                    If Val(ASCDATA1.GetDataValue) <> 0 Then
                        EMsg &= vbCr & "Some Shipments Selected are Not In Pick"
                    End If
                End If

                If SQLA("PICK_BATCH_NO") <> "" Then
                    ASCMAIN1.sql = "Select Count (*) from SOTPICK0 where NVL(PICK_BATCH_STATUS,'?') <> 'O'" & SQL_in("PICK_BATCH_NO")
                    If Val(ASCDATA1.GetDataValue) <> 0 Then
                        EMsg &= vbCr & "Some Pick Batches Selected are Not In Open"
                    End If
                    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" & SQL_in("PICK_BATCH_NO")
                    If Val(ASCDATA1.GetDataValue) <> 0 Then
                        EMsg &= vbCr & "Some Shipments in the Pick Batch Selected are Not In Pick"
                    End If
                End If

                ' Check LP_STATUS on any SHIP_BOL_NO that is touched by this De-Release

                'If SQLA("PICK_NO") <> "" Then
                '    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" _
                '        & " and SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")"
                '    If Val(ASCDATA1.GetDataValue) <> 0 Then
                '        EMsg &= vbCr & "Some Shipments Selected have been Transmitted to the 3PL"
                '    End If
                'End If

                'If SQLA("SHIP_BOL_NO") <> "" Then
                '    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" & SQL_in("SHIP_BOL_NO")
                '    If Val(ASCDATA1.GetDataValue) <> 0 Then
                '        EMsg &= vbCr & "Some Shipments Selected have been Transmitted to the 3PL"
                '    End If
                'End If

                'If SQLA("PICK_BATCH_NO") <> "" Then
                '    ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" & SQL_in("PICK_BATCH_NO")
                '    If Val(ASCDATA1.GetDataValue) <> 0 Then
                '        EMsg &= vbCr & "Some Shipments in the Pick Batch Selected have been Transmitted to the 3PL"
                '    End If
                'End If

                If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
                    Stop
                    If chkDeDeRelease.Checked Then
                        Stop
                        dederelease_PTs = True
                        EMsg = ""
                    End If
                End If
        End Select

    End Sub

    Overrides Sub Update_Record()
        If dederelease_PTs Then
            If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
                Stop
                DeDeRelease()
            End If
        Else
            DeRelease()
        End If

    End Sub

    Sub DeRelease()

        Dim sql_pick_D As String = " and SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTPICK1 & ")"

        ' Update Pick Ticket, Shipment Control & Carton Tables

        ' BeginTrans()

        ASCMAIN1.Progress("Now De-Releasing Pick Tickets", "")

        If SQLA("PICK_BATCH_NO") <> "" Then
            ASCMAIN1.sql = "Update SOTPICK0 Set " & vbCrLf _
                 & "  PICK_BATCH_STATUS = 'D'" & vbCrLf _
                 & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                 & ", LAST_DATE = SYSDATE" & vbCrLf _
                 & " where PICK_BATCH_NO in (" & SQLA("PICK_BATCH_NO", , True) & ")"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.Progress("-", "Status")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY,0)) QTY" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_CANC_REL,0)) QTY_CANC" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_BACK_REL,0)) QTY_BACK" & vbCrLf _
            & " from SOTORDR2,SOTPICK2,SOTPICK1,SOTORDR1 " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & sql_pick_D & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTSTAT2 " & vbCrLf _
            & " Set WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - R1.QTY, " & vbCrLf _
            & "     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) + R1.QTY + R1.QTY_CANC" & vbCrLf _
            & " where ITEM_CODE = R1.ITEM_CODE" & vbCrLf _
            & "   and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "   Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_OPEN)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.WHSE_CODE, -1 * R1.QTY, R1.QTY + R1.QTY_CANC);" & vbCrLf _
            & " End If;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.Progress("-", "Cartons")
        'ASCMAIN1.sql = "Delete FROM SOTCART2" & vbCrLf _
        '    & " where CART_NO in (Select CART_NO from SOTCART1 where PICK_NO in" & vbCrLf _
        '    & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
        '    & "))"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Delete from SOTCART1 where PICK_NO in" & vbCrLf _
        '    & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
        '    & ")"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Status")
        ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'O', ORDR_HOLD = '1', ORDR_HOLD_REASON = 'DE-RELEASED'" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Orders")
        ASCMAIN1.sql = "Update SOTORDR1 Set " _
            & "  ORDR_DATE_CLOSED = Null, ORDR_YYYYPP_CLOSED = Null, ORDR_YYYYPP_UPDATED = Null" & vbCrLf _
            & ", REORD_MEMO_IND = Null, ORDR_DATE_REL = Null, ORDR_REL_BATCH_NO = Null, ORDR_BATCHED = Null" & vbCrLf _
            & ", ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'Q'" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC,EVENT_KEY) " & vbCrLf _
            & " Select 'SOTORDR1',ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'DREL', 'Pick Ticket De-Released', PICK_NO" & vbCrLf _
            & " from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'O', ALLO_CTL_NO_REL = Null, ORDR_RELEASE = Null" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ' I DON'T KNOW WHY THIS NEXT UPDATE WAS SET UP AS A SEPARATE SQL STMT, 
        ' AND NOT COMBINED WITH THE ONE ABOVE
        ' SO I COMBINED IT
        'ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_RELEASE = Null" & vbCrLf _
        '    & " where ORDR_NO in" & vbCrLf _
        '    & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
        '    & ")"
        'ASCDATA1.ExecuteSQL()

        If ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "Delete from SOTCONF2" & vbCrLf _
            & " where ORDR_NO in " & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.Progress("-", "Tickets")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTPICK2.* from SOTPICK1,SOTPICK2" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & sql_pick_D & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update SOTORDR2 " & vbCrLf _
            & " Set ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - NVL(R1.PICK_QTY,0)," & vbCrLf _
            & "     ORDR_QTY_OPEN = NVL(ORDR_QTY_OPEN,0) + NVL(R1.PICK_QTY,0) + NVL(R1.PICK_QTY_CANC_REL,0)," & vbCrLf _
            & "     ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = R1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select DISTINCT ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTPICK1" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & ");" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTPICK1 set " & vbCrLf _
            & " PICK_STATUS = 'D', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Shipments")
        ASCMAIN1.sql = "Select SHIP_BOL_NO" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'P',1,0)) PICK" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'F',1,0)) SHIP" & vbCrLf _
            & ", Count (*) TOTAL" & vbCrLf _
            & " from SOTPICK1 " & vbCrLf _
            & " where SHIP_BOL_NO in " & vbCrLf _
            & " (Select SHIP_BOL_NO from " & SOTPICK1 & ")" & vbCrLf _
            & " group by SHIP_BOL_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)

            If Val(row.Item("PICK") & "") = 0 Then
                Dim SHIP_STATUS As String = ""
                If Val(row.Item("SHIP") & "") = 0 Then
                    SHIP_STATUS = "D"
                Else
                    SHIP_STATUS = "F" ' SHOULDNT SET F WITHOUT OTHER FIELDS WHICH GET THEIR VALUE VIA DATA ENTRY IN SHIPMENTS CONF
                    Stop ' MUST RESEARCH HOW THIS IS POSSIBLE, IF IT EVER HAPPENS
                    SHIP_STATUS = ""
                End If
                If SHIP_STATUS <> "" Then
                    ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_STATUS = '" & SHIP_STATUS & "'" _
                        & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                        & ", LP_STATUS = NULL" _
                        & " where SHIP_BOL_NO = '" & row.Item("SHIP_BOL_NO") & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        Next
    End Sub

    Sub DeDeRelease()

        ' do not do a batch - ie, SQLA("PICK_BATCH_NO") = ""
        ' just create a list of pick tickets in temp table SOTPICK1
        ' all pick tickets should have a status of D and all orders should have a status of O

        'SELECT SOTORDR1.CUST_CODE, SOTPICK1.PICK_BATCH_NO, SOTORDR1.ORDR_STATUS, SOTPICK1.PICK_STATUS, COUNT (*) PICKS
        ', MIN (SOTPICK1.PICK_NO) PMIN, MAX (SOTPICK1.PICK_NO) PMAX, MIN (SOTPICK1.ORDR_NO) OMIN, MAX(SOTPICK1.ORDR_NO) OMAX
        'FROM SOTPICK1,SOTORDR1 WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
        'AND SOTPICK1.PICK_BATCH_NO IN ('082494','082495','082496','082500','082520','082598')
        'GROUP BY SOTORDR1.CUST_CODE, SOTPICK1.PICK_BATCH_NO, SOTORDR1.ORDR_STATUS, SOTPICK1.PICK_STATUS
        'ORDER BY SOTORDR1.CUST_CODE, SOTPICK1.PICK_BATCH_NO, SOTORDR1.ORDR_STATUS, SOTPICK1.PICK_STATUS

        'CUST_CODE  PICK_B O P PICKS                  PMIN       PMAX       OMIN       OMAX      
        '---------- ------ - - ---------------------- ---------- ---------- ---------- ----------
        'MACYS      082494 D D 67                     0008055258 0008055324 0008126163 0008126285 < not to be included - order status is D
        'MACYS      082494 O D 66                     0008055192 0008055257 0008125636 0008125756
        'MACYS      082495 D D 95                     0008055419 0008055513 0008126180 0008126330 < not to be included - order status is D
        'MACYS      082495 O D 94                     0008055325 0008055418 0008125652 0008125801
        'MACYS      082496 D D 5                      0008055519 0008055523 0008126222 0008126235 < not to be included - order status is D
        'MACYS      082496 O D 5                      0008055514 0008055518 0008125694 0008125707
        'MACYS      082500 O D 43                     0008055538 0008055580 0008125606 0008125907
        'MACYS      082520 O D 141                    0008056118 0008056258 0008125476 0008125984
        'MACYS      082598 O D 128                    0008061141 0008061268 0008125461 0008125854

        'CREATE TABLE SOTPICK1_DDR AS
        'SELECT SOTPICK1.*
        'FROM SOTPICK1,SOTORDR1 WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
        'AND SOTPICK1.PICK_BATCH_NO IN ('082494','082495','082496','082500','082520','082598')
        'AND SOTORDR1.ORDR_STATUS = 'O'

        '08/28 
        ' RENAMED TO SOTPICK1_DDR_MACYS SO WE CAN USE SAME ROUTINE FOR DILLARDS (BELOW - 78 PTS)
        'CREATE TABLE SOTPICK1_DDR AS 
        'SELECT * FROM SOTPICK1 WHERE ORDR_NO IN (
        'SELECT  ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '0000133589')
        'AND PICK_STATUS ='D'

        'RENAME SOTPICK1_DDR TO SOTPICK1_DDR_MACYS2;
        'CREATE TABLE SOTPICK1_DDR AS 
        'SELECT * FROM SOTPICK1 WHERE ORDR_NO IN (
        'SELECT  ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO IN ( '0000134456','0000134436'))
        'AND PICK_STATUS ='D';

        ' in the above example use the records loaded into SOTPICK1_DDR


        Stop

        Dim sql_pick_D As String = " and SOTPICK1.PICK_NO in (Select PICK_NO from " & SOTPICK1 & ")"

        ASCMAIN1.Progress("Now De-De-Releasing Pick Tickets", "")

        'If SQLA("PICK_BATCH_NO") <> "" Then
        '    ASCMAIN1.sql = "Update SOTPICK0 Set " & vbCrLf _
        '         & "  PICK_BATCH_STATUS = 'D'" & vbCrLf _
        '         & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
        '         & ", LAST_DATE = SYSDATE" & vbCrLf _
        '         & " where PICK_BATCH_NO in (" & SQLA("PICK_BATCH_NO", , True) & ")"
        '    ASCDATA1.ExecuteSQL()
        'End If

        ASCMAIN1.Progress("-", "Status")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY,0)) QTY" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_CANC_REL,0)) QTY_CANC" & vbCrLf _
            & ", SUM (NVL(SOTPICK2.PICK_QTY_BACK_REL,0)) QTY_BACK" & vbCrLf _
            & " from SOTORDR2,SOTPICK2,SOTPICK1,SOTORDR1 " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & sql_pick_D & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update ICTSTAT2 " & vbCrLf _
            & " Set WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) + R1.QTY, " & vbCrLf _
            & "     WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - R1.QTY - R1.QTY_CANC" & vbCrLf _
            & " where ITEM_CODE = R1.ITEM_CODE" & vbCrLf _
            & "   and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & " If SQL%NOTFOUND Then" & vbCrLf _
            & "   Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_OPEN)" & vbCrLf _
            & "   Values (R1.ITEM_CODE, R1.WHSE_CODE, R1.QTY, -1 * (R1.QTY + R1.QTY_CANC));" & vbCrLf _
            & " End If;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Status")
        ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_STATUS = 'P', ORDR_HOLD = NULL, ORDR_HOLD_REASON = NULL" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Orders")
        ASCMAIN1.sql = "Update SOTORDR1 Set " _
            & "  ORDR_DATE_CLOSED = Null, ORDR_YYYYPP_CLOSED = Null, ORDR_YYYYPP_UPDATED = Null" & vbCrLf _
            & ", REORD_MEMO_IND = Null, ORDR_DATE_REL = Null, ORDR_REL_BATCH_NO = Null, ORDR_BATCHED = Null" & vbCrLf _
            & ", ORDR_REL_HOLD_CODES = NULL" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Status")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & "Select SOTPICK0.* from SOTPICK0" & vbCrLf _
            & " where PICK_BATCH_NO in" & vbCrLf _
            & " (Select PICK_BATCH_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & ");" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update SOTORDR1 Set " _
            & "  ORDR_DATE_REL = TRUNC(R1.INIT_DATE), ORDR_REL_BATCH_NO = Null" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & " and PICK_BATCH_NO = R1.PICK_BATCH_NO);" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()



        ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC,EVENT_KEY) " & vbCrLf _
            & " Select 'SOTORDR1',ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'DREL', 'Pick Ticket De-De-Released', PICK_NO" & vbCrLf _
            & " from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'P', ALLO_CTL_NO_REL = ALLO_CTL_NO, ORDR_RELEASE = Null" & vbCrLf _
            & " where ORDR_NO in" & vbCrLf _
            & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Tickets")
        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select SOTPICK2.* from SOTPICK1,SOTPICK2" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & sql_pick_D & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " Update SOTORDR2 " & vbCrLf _
            & " Set ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) + NVL(R1.PICK_QTY,0)," & vbCrLf _
            & "     ORDR_QTY_OPEN = NVL(ORDR_QTY_OPEN,0) - NVL(R1.PICK_QTY,0) - NVL(R1.PICK_QTY_CANC_REL,0)," & vbCrLf _
            & "     ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = R1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTPICK1 set " & vbCrLf _
            & " PICK_STATUS = 'P', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Shipments")
        ASCMAIN1.sql = "Select SHIP_BOL_NO" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'P',1,0)) PICK" & vbCrLf _
            & ", Sum (Decode (PICK_STATUS,'F',1,0)) SHIP" & vbCrLf _
            & ", Count (*) TOTAL" & vbCrLf _
            & " from SOTPICK1 " & vbCrLf _
            & " where SHIP_BOL_NO in " & vbCrLf _
            & " (Select SHIP_BOL_NO from " & SOTPICK1 & ")" & vbCrLf _
            & " group by SHIP_BOL_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)

            If Val(row.Item("PICK") & "") <> 0 Then
                Dim SHIP_STATUS As String = "P"
                If SHIP_STATUS <> "" Then
                    ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_STATUS = '" & SHIP_STATUS & "'" _
                        & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                        & ", LP_STATUS = '1'" _
                        & " where SHIP_BOL_NO = '" & row.Item("SHIP_BOL_NO") & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            End If
        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select DISTINCT ORDR_GROUP_NO from SOTSHIP1 " & vbCrLf _
            & " where SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTPICK1" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_pick_D) & ");" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Function Set_filter_for_Pick_Tickets_to_De_Release() As String

        Dim sql_pick As String = ""
        sql_pick &= SQL_in("PICK_BATCH_NO", "SOTPICK1.PICK_BATCH_NO")
        sql_pick &= SQL_in("SHIP_BOL_NO", "SOTPICK1.SHIP_BOL_NO")
        sql_pick &= SQL_in("PICK_NO", "SOTPICK1.PICK_NO")

        'sql_pick_all = sql_pick
        sql_pick = sql_pick & " and (SOTPICK1.PICK_STATUS = 'P' OR SOTPICK1.PICK_STATUS = 'C')"
     
        Return sql_pick
    End Function
End Class