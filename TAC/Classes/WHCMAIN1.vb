Public Class WHCMAIN1

    Public Shared Function Get_LP_XNO(MENU_ITEM_OBJECT As String, LP_XNO_RECORDS As Int64) As String
        Dim LP_XNO As String = ASCMAIN1.Next_Control_No("LP_XNO")

        ASCMAIN1.sql = "Insert into WHTLPXN1 (LP_XNO,LP_XNO_SOURCE,LP_XNO_RECORDS,LP_XNO_NOTES,INIT_OPER,INIT_DATE) " _
            & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,SYSDATE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVNVV", New Object() {LP_XNO, MENU_ITEM_OBJECT, LP_XNO_RECORDS, "", ASCMAIN1.USER_ID})
        Return LP_XNO
    End Function

    Public Shared Function Prepare_WHTSTYLX( _
        ByRef WHTSTYLX As String, _
        LP_CODE As String, _
        Optional initialize As Boolean = False, _
        Optional ICTSTAT2 As String = "ICTSTAT2") As String

        ASCMAIN1.sql = "((" _
        & "Select " & vbCrLf _
        & "ICTSTAT2.STYLE_CODE || ICTSTAT2.COLOR_CODE ITEM_CODE," & vbCrLf _
        & "ICTSTYL1.STYLE_DESC || ',' || ICTCOLR1.COLOR_DESC ITEM_DESC," & vbCrLf _
        & "NULL PPK_CODE," & vbCrLf _
        & "ICTSTAT2.STYLE_CODE," & vbCrLf _
        & "ICTSTAT2.COLOR_CODE," & vbCrLf _
        & "ICTSTYL1.STYLE_DESC," & vbCrLf _
        & "ICTCOLR1.COLOR_DESC," & vbCrLf _
        & "ICTSTYL1.INNER_PACK_QTY," & vbCrLf _
        & "ICTSTYL1.CARTON_PACK_QTY," & vbCrLf _
        & "ICTSTYL1.INIT_OPER,ICTSTYL1.INIT_DATE,ICTSTYL1.LAST_OPER,ICTSTYL1.LAST_DATE," & vbCrLf _
        & "'S' ITEM_TYPE" & vbCrLf _
        & " from ICTSTYL1, " & ICTSTAT2 & " ICTSTAT2, ICTCOLR1, ICTWHSE1" & vbCrLf _
        & " where ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE" & vbCrLf _
        & "   and ICTCOLR1.COLOR_CODE = ICTSTAT2.COLOR_CODE" & vbCrLf _
        & "   and ICTSTAT2.WHSE_CODE = ICTWHSE1.WHSE_CODE" & vbCrLf _
        & "   and ICTWHSE1.LP_CODE = '" & LP_CODE & "'" & vbCrLf _
        & ") union (" & vbCrLf _
        & "Select " & vbCrLf _
        & "ICTSTYC1.STYLE_CODE || ICTSTYC1.COLOR_CODE ITEM_CODE," & vbCrLf _
        & "ICTSTYL1.STYLE_DESC || ',' || ICTCOLR1.COLOR_DESC ITEM_DESC," & vbCrLf _
        & "NULL PPK_CODE," & vbCrLf _
        & "ICTSTYC1.STYLE_CODE," & vbCrLf _
        & "ICTSTYC1.COLOR_CODE," & vbCrLf _
        & "ICTSTYL1.STYLE_DESC," & vbCrLf _
        & "ICTCOLR1.COLOR_DESC," & vbCrLf _
        & "ICTSTYL1.INNER_PACK_QTY," & vbCrLf _
        & "ICTSTYL1.CARTON_PACK_QTY," & vbCrLf _
        & "ICTSTYL1.INIT_OPER,ICTSTYL1.INIT_DATE,ICTSTYL1.LAST_OPER,ICTSTYL1.LAST_DATE," & vbCrLf _
        & "'S' ITEM_TYPE" & vbCrLf _
        & " from ICTSTYL1, ICTSTYC1, ICTCOLR1" & vbCrLf _
        & " where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & vbCrLf _
        & "   and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" & vbCrLf _
        & "   and ICTSTYL1.LAST_DATE > SYSDATE -7" & vbCrLf _
        & ")" & vbCrLf _
        & " minus " & vbCrLf _
        & "Select " & vbCrLf _
        & "WHTSTYLX.ITEM_CODE," & vbCrLf _
        & "WHTSTYLX.ITEM_DESC," & vbCrLf _
        & "WHTSTYLX.PPK_CODE," & vbCrLf _
        & "WHTSTYLX.STYLE_CODE," & vbCrLf _
        & "WHTSTYLX.COLOR_CODE," & vbCrLf _
        & "WHTSTYLX.STYLE_DESC," & vbCrLf _
        & "WHTSTYLX.COLOR_DESC," & vbCrLf _
        & "WHTSTYLX.INNER_PACK_QTY," & vbCrLf _
        & "WHTSTYLX.CARTON_PACK_QTY," & vbCrLf _
        & "WHTSTYLX.INIT_OPER,WHTSTYLX.INIT_DATE,WHTSTYLX.LAST_OPER,WHTSTYLX.LAST_DATE," & vbCrLf _
        & "WHTSTYLX.ITEM_TYPE" & vbCrLf _
        & " from WHTSTYLX" & vbCrLf _
        & " where WHTSTYLX.LP_CODE = '" & LP_CODE & "'" & vbCrLf _
        & "   and WHTSTYLX.ITEM_TYPE = 'S'" & vbCrLf _
        & ") union (" & vbCrLf _
        & "Select " & vbCrLf _
        & "WHTPPKM1.PPK_CODE ITEM_CODE," & vbCrLf _
        & "WHTPPKM1.PPK_DESC ITEM_DESC," & vbCrLf _
        & "WHTPPKM1.PPK_CODE," & vbCrLf _
        & "NULL STYLE_CODE," & vbCrLf _
        & "NULL COLOR_CODE," & vbCrLf _
        & "NULL STYLE_DESC," & vbCrLf _
        & "NULL COLOR_DESC," & vbCrLf _
        & "NULL INNER_PACK_QTY," & vbCrLf _
        & "NULL CARTON_PACK_QTY," & vbCrLf _
        & "WHTPPKM1.INIT_OPER,WHTPPKM1.INIT_DATE,WHTPPKM1.LAST_OPER,WHTPPKM1.LAST_DATE," & vbCrLf _
        & "'P' ITEM_TYPE" & vbCrLf _
        & " from WHTPPKM1" & vbCrLf _
        & " minus " & vbCrLf _
        & "Select " & vbCrLf _
        & "WHTSTYLX.ITEM_CODE," & vbCrLf _
        & "WHTSTYLX.ITEM_DESC," & vbCrLf _
        & "WHTSTYLX.PPK_CODE," & vbCrLf _
        & "WHTSTYLX.STYLE_CODE," & vbCrLf _
        & "WHTSTYLX.COLOR_CODE," & vbCrLf _
        & "WHTSTYLX.STYLE_DESC," & vbCrLf _
        & "WHTSTYLX.COLOR_DESC," & vbCrLf _
        & "WHTSTYLX.INNER_PACK_QTY," & vbCrLf _
        & "WHTSTYLX.CARTON_PACK_QTY," & vbCrLf _
        & "WHTSTYLX.INIT_OPER,WHTSTYLX.INIT_DATE,WHTSTYLX.LAST_OPER,WHTSTYLX.LAST_DATE," & vbCrLf _
        & "WHTSTYLX.ITEM_TYPE" & vbCrLf _
        & " from WHTSTYLX" & vbCrLf _
        & " where WHTSTYLX.LP_CODE = '" & LP_CODE & "'" & vbCrLf _
        & "   and WHTSTYLX.ITEM_TYPE = 'P'" & vbCrLf _
        & ")"

        If initialize Then
            WHTSTYLX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & WHTSTYLX & " Add Primary Key (ITEM_CODE)")
        Else
            ASCMAIN1.sql = "Insert into " & WHTSTYLX & " " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from WHTSTYLX where LP_CODE = '" & LP_CODE & "'" _
            & " and (ITEM_CODE) in (Select ITEM_CODE from " & WHTSTYLX & ")"
            ASCDATA1.ExecuteSQL()


            ASCMAIN1.sql = "Insert into WHTSTYLX " _
                & "(LP_CODE, ITEM_CODE, ITEM_DESC, PPK_CODE" _
                & ", STYLE_CODE, COLOR_CODE, STYLE_DESC, COLOR_DESC" _
                & ", INNER_PACK_QTY, CARTON_PACK_QTY" _
                & ", INIT_OPER, INIT_DATE, LAST_OPER, LAST_DATE" _
                & ", ITEM_TYPE, STATUS) " _
                & "Select '" & LP_CODE & "' LP_CODE, " & WHTSTYLX & ".*,'0' STATUS from " & WHTSTYLX
            ASCDATA1.ExecuteSQL()
        End If

        Return WHTSTYLX
    End Function

    Public Shared Sub Send_WHTSTYLX(LP_CODE As String, LP_XNO As String)

        ASCMAIN1.sql = "Update WHTSTYLX Set STATUS = '1',LP_XNO = '" & LP_XNO & "'" _
            & " where LP_CODE = '" & LP_CODE & "' and STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ADS.WHTSTYLX@ADSIIS where LP_CODE = '" & LP_CODE & "'" _
            & " and (ITEM_CODE) in (Select ITEM_CODE from WHTSTYLX where LP_XNO = '" & LP_XNO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ADS.WHTPPKM1@ADSIIS where " _
            & " PPK_CODE in (Select ITEM_CODE from WHTSTYLX where ITEM_TYPE = 'P' and LP_XNO = '" & LP_XNO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ADS.WHTPPKM2@ADSIIS where " _
            & " PPK_CODE in (Select ITEM_CODE from WHTSTYLX where ITEM_TYPE = 'P' and LP_XNO = '" & LP_XNO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ADS.WHTPPKM1@ADSIIS" _
            & " Select * from WHTPPKM1 where PPK_CODE in " _
            & "(Select ITEM_CODE from WHTSTYLX where ITEM_TYPE = 'P' and LP_XNO = '" & LP_XNO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ADS.WHTPPKM2@ADSIIS" _
            & " Select * from WHTPPKM2 where PPK_CODE in " _
            & "(Select ITEM_CODE from WHTSTYLX where ITEM_TYPE = 'P' and LP_XNO = '" & LP_XNO & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ADS.WHTSTYLX@ADSIIS" _
            & " Select * from WHTSTYLX where LP_CODE = '" & LP_CODE & "' and LP_XNO = '" & LP_XNO & "'"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Shared Sub UpdateADSAndImport()
        If Not ASCMAIN1.Running_in_VS AndAlso Not ASCMAIN1.DBS_SERVER = "" Then

            ASCMAIN1.sql = "UPDATE ADS.INVADJ@ADSIIS SET STATUS = 'V' WHERE STATUS = '0'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO WHTIADJ1 (TRANS_SEQ, LP_CODE,WHSE_CODE,STATUS" & vbCrLf _
                & ",TRNDTE,ITEM_CODE,ADJQTY,REACOD,ADJ_REF1,ADJ_REF2,ABS_STATUS)" & vbCrLf _
                & "SELECT TRANS_SEQ,LP_CODE,WHSE_CODE,STATUS,TRNDTE,ITEM_CODE," & vbCrLf _
                & "ADJQTY, REACOD,ADJ_REF1,ADJ_REF2,'N' FROM ADS.INVADJ@ADSIIS WHERE STATUS = 'V'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "UPDATE ADS.INVADJ@ADSIIS SET STATUS = '1' WHERE STATUS = 'V'"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Public Shared Sub Update_ADS_SOTSHIP1()
        If ASCMAIN1.Running_in_VS Then Stop
        ASCMAIN1.sql = "Update WHTPARM1 Set WH_PARM_ADS_LAST_UPDATE = SYSDATE where WH_PARM_ADS_LAST_UPDATE < SYSDATE - 1/(24 * 4)"
        Dim I As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        If I = 1 Then
            ASCMAIN1.sql = "BEGIN" & vbCrLf _
                & "DELETE FROM ADS.SOTSHIP0_3PL@ADSIIS;" & vbCrLf _
                & "INSERT INTO ADS.SOTSHIP0_3PL@ADSIIS SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE SHIP_STATUS = 'P';" & vbCrLf _
                & "BEGIN DECLARE CURSOR C1 IS" & vbCrLf _
                & "SELECT SHIP_BOL_NO, SHIP_NOTES, SHIP_LOAD_NO, SHIP_APPT_NO, SHIP_DATE_PLANNED, LP_STATUS_TS_3PL" & vbCrLf _
                & " FROM ADS.SOTSHIP1_3PL@ADSIIS WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM ADS.SOTSHIP0_3PL@ADSIIS)" & vbCrLf _
                & "AND (SHIP_NOTES IS NOT NULL OR SHIP_LOAD_NO IS NOT NULL OR SHIP_APPT_NO IS NOT NULL);" & vbCrLf _
                & "rowSOTSHIP1 SOTSHIP1%ROWTYPE;" & vbCrLf _
                & "BEGIN" & vbCrLf _
                & "FOR R1 IN C1 LOOP" & vbCrLf _
                & "SELECT * INTO ROWSOTSHIP1 FROM SOTSHIP1 WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "UPDATE SOTSHIP1 SET SHIP_NOTES_3PL = R1.SHIP_NOTES, SHIP_LOAD_NO = R1.SHIP_LOAD_NO, " & vbCrLf _
                & "SHIP_APPT_NO = R1.SHIP_APPT_NO, SHIP_DATE_PLANNED = R1.SHIP_DATE_PLANNED" & vbCrLf _
                & "WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "IF NVL(R1.SHIP_NOTES,'?') <> NVL(ROWSOTSHIP1.SHIP_NOTES_3PL,'?') THEN" & vbCrLf _
                & "INSERT INTO ASTAUDT1 (TABLE_NAME,KEY_VALUE,COLUMN_NAME,USER_ID,INIT_DATE,OLD_VALUE,NEW_VALUE)" & vbCrLf _
                & "VALUES ('SOTSHIP1',R1.SHIP_BOL_NO,'SHIP_NOTES_3PL','auto',SYSDATE,ROWSOTSHIP1.SHIP_NOTES_3PL,R1.SHIP_NOTES);" & vbCrLf _
                & "END IF;" & vbCrLf _
                & "END LOOP; END; END;" & vbCrLf _
                & "END;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, True)
        End If
    End Sub

    Public Shared Sub Update_WHTLOCBX(  frm As ASFBASE0)

        If frm.dst.Tables.Contains("WHTLOCB1") Then
            frm.dst.Tables("WHTLOCB1").Rows.Clear()
            frm.dst.Tables("WHTLOCB2").Rows.Clear()
        Else
            frm.Create_TDA(frm.dst.Tables.Add, "WHTLOCB1", "*")
            frm.Create_TDA(frm.dst.Tables.Add, "WHTLOCB2", "*")
        End If

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            Dim WHSE_TRAN_NO As String = row.Item("WHSE_TRAN_NO")
            Dim WHSE_TRAN_LNO As Integer = Val(row.Item("WHSE_TRAN_LNO") & "")
            Dim WHSE_TRAN_TYPE As String = row.Item("WHSE_TRAN_TYPE")
            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
            Dim BAR_CODE As String = "0000000000" ' row.Item("BAR_CODE")
            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE") & ""

            If LOCATION_CODE.Length = 0 Then LOCATION_CODE = "000000"

            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim WHSE_TRAN_QTY As Int64 = Val(row.Item("WHSE_TRAN_QTY") & "")

            Dim rowWHTLOCB1 As DataRow = frm.dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE})
            If rowWHTLOCB1 Is Nothing Then
                frm.Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE}, False)
                rowWHTLOCB1 = frm.dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, ITEM_CODE})
            End If

            If rowWHTLOCB1 Is Nothing Then
                rowWHTLOCB1 = frm.dst.Tables("WHTLOCB1").NewRow
                With rowWHTLOCB1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOCATION_CODE") = LOCATION_CODE
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("LOCATION_QTY") = WHSE_TRAN_QTY
                End With
                frm.dst.Tables("WHTLOCB1").Rows.Add(rowWHTLOCB1)
            Else
                rowWHTLOCB1.Item("LOCATION_QTY") = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "") + WHSE_TRAN_QTY
            End If

            Dim rowWHTLOCB2 As DataRow = frm.dst.Tables("WHTLOCB2").NewRow
            With rowWHTLOCB2
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("BAR_CODE") = BAR_CODE
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("WHSE_TRAN_QTY") = WHSE_TRAN_QTY
                .Item("WHSE_TRAN_TYPE") = WHSE_TRAN_TYPE
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                .Item("INIT_DATE") = frm.DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE_OTHER") = ""
                .Item("SESSION_NO") = ""
            End With
            frm.dst.Tables("WHTLOCB2").Rows.Add(rowWHTLOCB2)
        Next

        frm.Update_Record_TDA("WHTLOCB1")
        frm.Update_Record_TDA("WHTLOCB2")

        frm.dst.Tables("WHTLOCB1").Rows.Clear()
        frm.dst.Tables("WHTLOCB2").Rows.Clear()
    End Sub

    Public Shared Sub Prepare_Carton_Data_3PL(SOTSHIPX As String, LP_CODE As String, ASW As Dictionary(Of String, String))

        'Stop
        ' this is a copy of code taken from WHFSPCK1 and we might not need this routine

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  SOTCART1.CART_NO" & vbCrLf _
            & ", SOTCART1.CART_FREIGHT" & vbCrLf _
            & ", SOTCART1.PICK_NO" & vbCrLf _
            & ", SOTCART1.CART_TOTAL_UNITS" & vbCrLf _
            & ", SOTCART1.CART_TOTAL_WGT_ACTUAL" & vbCrLf _
            & ", SOTCART1.CART_TOTAL_WGT_CALC" & vbCrLf _
            & ", SOTCART1.CART_TRACKING_NO" & vbCrLf _
            & ", SOTCART1.CART_SEQ" & vbCrLf _
            & ", SOTCART1.CART_MEMO" & vbCrLf _
            & ", SOTCART1.CART_TYPE" & vbCrLf _
            & " from SOTCART1,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and (NVL(SOTSHIPX.EDI856,'0') = '1' or NVL(SOTSHIPX.SHIP_CART_REQD,'0') = '1')" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO"
        ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART1_3PL") & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  MIN (SOTCART1.CART_NO) CART_NO" & vbCrLf _
            & ", SUM (SOTCART1.CART_FREIGHT) CART_FREIGHT" & vbCrLf _
            & ", SOTCART1.PICK_NO" & vbCrLf _
            & ", SUM (SOTCART1.CART_TOTAL_UNITS) CART_TOTAL_UNITS" & vbCrLf _
            & ", SUM (SOTCART1.CART_TOTAL_WGT_ACTUAL) CART_TOTAL_WGT_ACTUAL" & vbCrLf _
            & ", SUM (SOTCART1.CART_TOTAL_WGT_CALC) CART_TOTAL_WGT_CALC" & vbCrLf _
            & ", NULL CART_TRACKING_NO" & vbCrLf _
            & ", 0 CART_SEQ" & vbCrLf _
            & ", NULL CART_MEMO" & vbCrLf _
            & ", NULL CART_TYPE" & vbCrLf _
            & " from SOTCART1,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and (NVL(SOTSHIPX.EDI856,'0') = '0' and NVL(SOTSHIPX.SHIP_CART_REQD,'0') <> '1')" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
            & " group by SOTCART1.PICK_NO"
        ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART1_3PL") & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  SOTCART2.CART_NO" & vbCrLf _
            & ", SOTCART2.CART_LNO" & vbCrLf _
            & ", SOTCART2.QTY_PACKED" & vbCrLf _
            & ", WHTSTYLX.ITEM_CODE" & vbCrLf _
            & ", 0 QTY_PACKED_ACT" & vbCrLf _
            & ", SOTCART1.PICK_NO" & vbCrLf _
            & ", SOTCART2.ORDR_LNO PICK_LNO" & vbCrLf _
            & " from SOTCART1,SOTCART2,SOTPICK1,SOTORDR2,WHTSTYLX," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO" & vbCrLf _
            & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
            & "   and WHTSTYLX.ITEM_CODE (+) = NVL(SOTORDR2.ITEM_CODE,SOTORDR2.STYLE_CODE || SOTORDR2.COLOR_CODE)" & vbCrLf _
            & "   and WHTSTYLX.LP_CODE (+) = '" & LP_CODE & "'" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO"
        ASCDATA1.ExecuteSQL("Insert into " & ASW("SOTCART2_3PL") & " " & ASCMAIN1.sql)

        ' Re-write all SOTCART2 records to use the single Carton that belongs to the Pick Ticket for ADS-non-Cartonized Orders
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select SOTCART1_3PL.CART_NO, SOTCART1_3PL.PICK_NO" & vbCrLf _
            & "   from SOTPICK1," & ASW("SOTCART1_3PL") & " SOTCART1_3PL," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & "   where SOTPICK1.PICK_NO = SOTCART1_3PL.PICK_NO" & vbCrLf _
            & "     and SOTSHIPX.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
            & "     and (NVL(SOTSHIPX.EDI856,'0') = '0' and NVL(SOTSHIPX.SHIP_CART_REQD,'0') <> '1');" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ASW("SOTCART2_3PL") & " SOTCART2_3PL" & vbCrLf _
            & "    Set CART_LNO = -1 * CART_LNO where PICK_NO = R1.PICK_NO;" & vbCrLf _
            & "   Insert into " & ASW("SOTCART2_3PL") & " SOTCART2_3PL" & vbCrLf _
            & "    Select R1.CART_NO CART_NO, ROWNUM CART_LNO, QTY_PACKED, ITEM_CODE, 0 QTY_PACKED_ACT, PICK_NO, PICK_LNO" & vbCrLf _
            & "     from (Select ITEM_CODE, PICK_NO, PICK_LNO, SUM (QTY_PACKED) QTY_PACKED" & vbCrLf _
            & "     from " & ASW("SOTCART2_3PL") & " SOTCART2_3PL where PICK_NO = R1.PICK_NO" & vbCrLf _
            & "     group by ITEM_CODE, PICK_NO, PICK_LNO);" & vbCrLf _
            & "   Delete from " & ASW("SOTCART2_3PL") & " SOTCART2_3PL" & vbCrLf _
            & "    where PICK_NO = R1.PICK_NO and CART_LNO < 0;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub
End Class