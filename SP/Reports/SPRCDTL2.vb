Public Class SPRCDTL2

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -60, 0, -18)
        Set_cmbYP("RYP1", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -60, 0, -12)
        Set_cmbYP("RYP2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -60, 0, -18)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        MyBase.Get_SQL("*")

        Dim startPeriod As String = cmbRYP0.SelectedRow.Cells(0).Value
        Dim endPeriod As String = cmbRYP1.SelectedRow.Cells(0).Value
        Dim salesPeriod As String = cmbRYP2.SelectedRow.Cells(0).Value

        Dim startDate As String = startPeriod.Substring(4, 2) & "/01/" & startPeriod.Substring(0, 4)
        Dim endDate As String = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(endPeriod.Substring(4, 2) & "/01/" & endPeriod.Substring(0, 4)))).ToString("MM/dd/yyyy")

        sql = "Select SPTCOOP1.*, SPTCOOP3.AUTH_LNO, SPTCOOP3.AUTH_SLNO, " & vbCr
        sql &= " SPTCOOP1.OPEN_AMT * (SPTCOOP3.DIST_AMT / (SPTCOOP1.QTY * SPTCOOP1.VEHICLE_CPM / 1000 + SPTCOOP1.OTHER_COST)) OPEN_AMT_COLLECTION," & vbCr
        sql &= " SPTCOOP1.PAID_AMT * (SPTCOOP3.DIST_AMT / (SPTCOOP1.QTY * SPTCOOP1.VEHICLE_CPM / 1000 + SPTCOOP1.OTHER_COST)) PAID_AMT_COLLECTION," & vbCr
        sql &= " SPTCOOP3.ITEM_CODE, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, ICTCOLL1.BRAND_CODE," & vbCr
        sql &= " 0 QTY_SHP, SYSDATE SHIP_DATE, SYSDATE CANC_DATE, 0 SALES, 0 QTY_OPN, 0 QTY_RSV"
        sql &= " from SPTCOOP1, SPTCOOP3, ICTCOLL1" & sql_TABLE_NAMEs.Replace(",SPTCOOP1", "").Replace(",SPTCOOP3", "").Replace(",ICTCOLL1", "") & vbCr
        sql &= " where SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO" & vbCr
        sql &= " and SPTCOOP1.OPS_YYYYPP between '" & startPeriod & "' and '" & endPeriod & "'" & vbCr
        sql &= " and (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000 + NVL(SPTCOOP1.OTHER_COST, 0)) <> 0"

        sql &= sql_JOIN
        sql &= sql_WHERE

        Dim wkTable As String = ASCMAIN1.Temp_Table(sql)
        sql = "UPDATE " & wkTable & " SET SHIP_DATE = NULL, CANC_DATE = NULL"
        ASCDATA1.ExecuteSQL(sql)

        Dim sqlShip As String = "Select SOTORDR1.CUST_CODE, SOTORDR2.ITEM_CODE" _
            & ", Sum (SOTORDR2.ORDR_QTY_SHIP) QTY, Max (SOTORDR1.ORDR_SHIP_DATE) SHIP, Max (SOTORDR1.ORDR_CANCEL_DATE) CANC" _
            & ", Sum (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) SALES" _
            & " from SOTORDR2, SOTORDR1 " _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & " and SOTORDR2.ORDR_STATUS = 'F' and SOTORDR2.ORDR_YYYYPP_UPDATED >= '" & salesPeriod & "'" _
            & " and (SOTORDR2.CUST_CODE, SOTORDR2.ITEM_CODE) IN (SELECT CUST_CODE, ITEM_CODE FROM " & wkTable & ")" _
            & " GROUP BY SOTORDR1.CUST_CODE, SOTORDR2.ITEM_CODE"

        'dynSHP.Refresh()
        'dynMEWCOOPX.Edit()
        'dynMEWCOOPX.Fields("QTY_SHP").Value = Val(dynSHP.Fields("QTY").Value & "")
        'dynMEWCOOPX.Fields("SHIP_DATE").Value = dynSHP.Fields("SHIP").Value
        'dynMEWCOOPX.Fields("CANC_DATE").Value = dynSHP.Fields("CANC").Value
        'dynMEWCOOPX.Fields("SALES").Value = Val(dynMEWCOOPX.Fields("SALES").Value & "") + Val(dynSHP.Fields("SALES").Value & "")

        sql = "BEGIN DECLARE CURSOR C1 IS " & sqlShip & ";"
        sql &= " BEGIN FOR R1 IN C1 LOOP"
        sql &= "    UPDATE " & wkTable & " SET QTY_SHP = NVL(R1.QTY, 0), SHIP_DATE = R1.SHIP, CANC_DATE = R1.CANC, SALES = SALES + NVL(R1.SALES, 0)"
        sql &= "        WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE;"
        sql &= " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        Dim sqlOpen As String = "Select SOTORDR1.CUST_CODE, SOTORDR2.ITEM_CODE" _
            & ", Sum (SOTORDR2.ORDR_QTY) QTY, Max (SOTORDR1.ORDR_SHIP_DATE) SHIP, Max (SOTORDR1.ORDR_CANCEL_DATE) CANC" _
            & ", Sum (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) SALES" _
            & " from SOTORDR2, SOTORDR1 " _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & " and (SOTORDR2.ORDR_STATUS = 'O' or SOTORDR2.ORDR_STATUS = 'P')" _
            & " and (SOTORDR2.CUST_CODE, SOTORDR2.ITEM_CODE) IN (SELECT CUST_CODE, ITEM_CODE FROM " & wkTable & ")" _
            & " GROUP BY SOTORDR1.CUST_CODE, SOTORDR2.ITEM_CODE"

        'dynMEWCOOPX.Fields("QTY_OPN").Value = Val(dynOPN.Fields("QTY").Value & "")
        'If IsNull(dynMEWCOOPX.Fields("SHIP_DATE").Value) Then
        '    dynMEWCOOPX.Fields("SHIP_DATE").Value = dynOPN.Fields("SHIP").Value
        'End If
        'If IsNull(dynMEWCOOPX.Fields("CANC_DATE").Value) Then
        '    dynMEWCOOPX.Fields("CANC_DATE").Value = dynOPN.Fields("CANC").Value
        'End If
        'dynMEWCOOPX.Fields("SALES").Value = Val(dynMEWCOOPX.Fields("SALES").Value & "") + Val(dynOPN.Fields("SALES").Value & "")

        ''        dynRSV.Refresh
        ''        dynMEWCOOPX.Fields("QTY_RSV").Value = Val(dynRSV.Fields("QTY").Value & "")
        'dynMEWCOOPX.Fields("QTY_RSV").Value = 0

        sql = "BEGIN DECLARE CURSOR C1 IS " & sqlOpen & ";"
        sql &= " BEGIN FOR R1 IN C1 LOOP"
        sql &= "    UPDATE " & wkTable & " SET QTY_OPN = NVL(R1.QTY, 0) WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE;"
        sql &= "    UPDATE " & wkTable & " SET SHIP_DATE = R1.SHIP WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE AND SHIP_DATE IS NULL;"
        sql &= "    UPDATE " & wkTable & " SET CANC_DATE = R1.CANC WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE AND CANC_DATE IS NULL;"
        sql &= "    UPDATE " & wkTable & " SET SALES = SALES + NVL(R1.SALES, 0) WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE;"
        sql &= " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        ' *****************************************************************************
        ' Need code for reserved quantities
        ' *****************************************************************************

        sql = "Select * from " & wkTable
        If Not dst.Tables.Contains("SPTCOOPX") Then
            Create_TDA(dst.Tables.Add, "SPTCOOPX", sql, 0, False, "", 0)
        End If
        Fill_Records("SPTCOOPX", String.Empty, True, sql)

        sql = "Select " & sql_SELECT_cols.Replace("SPTCOOP1.", "").Replace("SPTCOOP3.", "").Replace("ICTCOLL1.", "") & vbCr
        sql &= ", AUTH_NO, AUTH_LNO, AUTH_SLNO, 0 TOTAL "
        sql &= " FROM " & wkTable

        sql = "Insert Into " & ASTSRPT1 & " " & sql
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)

    End Sub

    Overrides Sub Build_Report_File_Post_Process()
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

End Class