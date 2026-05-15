Public Class SORSLSB1
    Dim SATBUDW1 As String
    ' NEED TO CREATE A RETURNS BUDGET FILE
    ' NEED TO GET THE SHAPE AND DIMENSIONS OF THE WHOLESALE SHIPMENTS BUDGET FILE
    ' W/S BUDGETS NEED COLLECTION AND BASIC/PROMO
    ' get column captions out of data and into fields of pivot table for G1, G2, G3

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
        'chkFLASH.Checked = False
        'chkFLASH.Visible = False

        'grpReportVersion.Visible = False
        'End If

        dteTODAY1.DateTime = Now.Date.AddDays(-1)
        dteTODAY2.DateTime = Now.Date.AddDays(-1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim TYP As String = ASCMAIN1.CYP
        Dim TYP01 As String = Mid(TYP, 1, 4) & "01"
        Dim TYP12 As String = Mid(TYP, 1, 4) & "12"
        Dim LYP As String = ASCMAIN1.Period_Calc(TYP, -12)
        Dim LYP01 As String = Mid(LYP, 1, 4) & "01"
        Dim LYP12 As String = Mid(LYP, 1, 4) & "12"

        SATBUDW1 = TAC.SOCMAIN1.Setup_Budgets_by_Customer

        ' Prepare Work Tables
        ASCMAIN1.sql = "" _
            & "Select ORDR_NO, ORDR_SHIP_DATE from SOTORDR1 where ORDR_STATUS in ('O','P') and ORDR_TYPE_CODE in ('REG','B2C')" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ORDR_NO, ORDR_SHIP_DATE from SOTORDR1 where ORDR_STATUS = 'C' and ORDR_TYPE_CODE in ('REG','B2C') and ORDR_YYYYPP_UPDATED = '" & TYP & "'" & vbCrLf _
            & " union " & vbCrLf _
            & "Select Distinct SOTINVH1.ORDR_NO, SOTORDR1.ORDR_SHIP_DATE from SOTINVH1,SOTORDR1 where SOTINVH1.ORDR_YYYYPP_UPDATED = '" & TYP & "' and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO and SOTORDR1.ORDR_TYPE_CODE in ('REG','B2C')"

        Dim SOTORDR1_FCOP As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1_FCOP & " Add Primary Key (ORDR_NO)")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim FX As String = ""
        Dim FACTOR As Decimal = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If
        If FACTOR <> 1 Then
            FX = "/" & CStr(FACTOR)
        End If

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        Dim sqlX12 As String = ""
        Dim sqlX As String = ""
        Dim sqlYP As String = ""
        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", TYP)
        Dim TYTM_DTE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim TYTM_DTEo As String = Format(TYTM_DTE, "dd-MMM-yyyy")
        rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(TYP, -1))
        Dim TYTM_DTS As Date = CDate(rowGLTPARM2.Item("PRD_END_DATE")).AddDays(1)
        Dim TYTM_DTSo As String = Format(TYTM_DTS, "dd-MMM-yyyy")

        Dim TYNM_DTS As Date = TYTM_DTE.AddDays(1)
        Dim TYNM_DTSo As String = Format(TYNM_DTS, "dd-MMM-yyyy")
        rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(TYP, 1))
        Dim TYNM_DTE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim TYNM_DTEo As String = Format(TYNM_DTE, "dd-MMM-yyyy")


        ASCMAIN1.Progress("Bookings")
        MyBase.Get_SQL("*")

        sqlX = "NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        Dim sqlXC As String = "NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        sqlYP = "SOTORDR1.ORDR_YYYYPP_BOOKED"

        sql_Data = "" _
            & ", SUM (CASE WHEN " & sqlYP & " < '" & TYP & "' AND SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) BOOKED_PRV" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' AND SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) BOOKED_CUR" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' AND SOTORDR1.ORDR_SHIP_DATE > '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) BOOKED_FUT" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "' THEN " & sqlXC & " ELSE 0 END) LOST_SALES_TYTM" & vbCrLf

        sql_Cols = "" _
            & ",BOOKED_PRV,BOOKED_CUR,BOOKED_FUT,LOST_SALES_TYTM"

        sql_filter = "" _
            & " and SOTORDR1.ORDR_SHIP_DATE >= '" & TYTM_DTSo & "'" & vbCrLf _
            & " and (SOTORDR1.ORDR_YYYYPP_UPDATED IS NULL OR SOTORDR1.ORDR_YYYYPP_UPDATED = '" & TYP & "')" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS <> 'D' and NVL(SOTORDR2.ORDR_RELEASE,'?') <> 'D' and SOTORDR1.ORDR_TYPE_CODE in ('REG','B2C')"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Bookings YTD")
        MyBase.Get_SQL("*")

        sqlX = "NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        'Dim sqlXC As String = "NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        sqlYP = "SOTORDR1.ORDR_YYYYPP_BOOKED"

        sql_Data = "" _
            & ", SUM (" & sqlX & ") BOOKED_TYTD" & vbCrLf

        sql_Cols = "" _
            & ",BOOKED_TYTD"

        sql_filter = "" _
            & " and SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "'" & vbCrLf _
            & " and " & sqlYP & " BETWEEN '" & TYP01 & "' and '" & TYP & "'" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS <> 'D' and NVL(SOTORDR2.ORDR_RELEASE,'?') <> 'D' and SOTORDR1.ORDR_TYPE_CODE in ('REG','B2C')"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Released & OTS")
        'MyBase.Get_SQL("*")

        sqlX = "(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        'sqlYP = "SOTORDR1.ORDR_SHIP_DATE"

        '            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE between '" & TYTM_DTSo & "' AND '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) OTS_M01" & vbCrLf _

        sql_Data = "" _
            & $", SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'P' and SOTORDR1.ORDR_SHIP_DATE <= '{TYTM_DTEo}' THEN NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) " & FX & " RELEASED_CUR" & vbCrLf _
            & $", SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'P' and SOTORDR1.ORDR_SHIP_DATE >  '{TYTM_DTEo}' THEN NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) " & FX & " RELEASED_FUT" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'P' THEN NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) " & FX & " RELEASED" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) OTS_M01" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE between '" & TYNM_DTSo & "' AND '" & TYNM_DTEo & "' THEN " & sqlX & " ELSE 0 END) OTS_M02" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE > '" & TYNM_DTEo & "' THEN " & sqlX & " ELSE 0 END) OTS_M03" & vbCrLf

        sql_Cols = "" _
            & ",RELEASED_CUR,RELEASED_FUT,RELEASED,OTS_M01,OTS_M02,OTS_M03"

        sql_filter = " and SOTORDR1.ORDR_STATUS in ('O','P') and SOTORDR1.ORDR_TYPE_CODE in ('REG','B2C')"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Carried Forward")
        'MyBase.Get_SQL("*")

        sqlX = "NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        'sqlYP = "SOTORDR1.ORDR_YYYYPP_BOOKED"

        sql_Data = "" _
            & ", SUM (" & sqlX & ") CARRIED_FWD" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & TYTM_DTEo & "' THEN " & sqlXC & " ELSE 0 END) LOST_SALES_TYTM" & vbCrLf

        sql_Cols = "" _
            & ",CARRIED_FWD, LOST_SALES_TYTM"

        sql_filter = "" _
            & " and SOTORDR1_FCOP.ORDR_SHIP_DATE < '" & TYTM_DTSo & "'" & vbCrLf _
            & " and SOTORDR1.ORDR_NO = SOTORDR1_FCOP.ORDR_NO" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS <> 'D' and SOTORDR1.ORDR_TYPE_CODE in ('REG','B2C')"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTORDR1" & "," & SOTORDR1_FCOP & " SOTORDR1_FCOP" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Shipments")
        MyBase.Get_SQL("I")

        sqlX = "NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)" & FX
        sqlYP = "SOTINVH1.ORDR_YYYYPP_UPDATED"
        sql_Data = "" _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & LYP & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_LYTM" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & LYP01 & "' AND " & sqlYP & " <= '" & LYP & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_LYTD" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_TYTM" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' AND  SOTORDR1.ORDR_SHIP_DATE > '" & TYTM_DTEo & "' THEN " & sqlX & " ELSE 0 END) BOOK_FUT_SHIP_CUR_MTD" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & TYP01 & "' AND " & sqlYP & " <= '" & TYP & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_TYTD" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & LYP01 & "' AND " & sqlYP & " <= '" & LYP12 & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_LYTOT" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' AND SOTINVH1.INV_DATE >= '" & Format(dteTODAY1.Value, "dd-MMM-yyyy") & "' AND SOTINVH1.INV_DATE <= '" & Format(dteTODAY2.Value, "dd-MMM-yyyy") & "' THEN " & sqlX & " ELSE 0 END) SHIPPED_TODAY" & vbCrLf

        ' KINDA THINK THAT LOST SALES MTD BELONGS HERE TOO INSTEAD OF UP NEAR BOOKINGS

        ' Dim sqlXC As String = "NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX


        sql_Cols = "" _
            & ",SHIPPED_LYTM,SHIPPED_LYTD,SHIPPED_TYTM,BOOK_FUT_SHIP_CUR_MTD,SHIPPED_TYTD,SHIPPED_LYTOT,SHIPPED_TODAY" ',LOST_SALES_TYTD"

        sql_filter = "" _
            & " and " & sqlYP & " BETWEEN '" & LYP01 & "' and '" & TYP & "'" _
            & " and SOTINVH2.INV_TYPE = 'I'" _
            & " and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTORDR1, SOTINVH1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Returns")
        MyBase.Get_SQL("I")

        sqlX = "-1 * NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)" & FX
        sqlYP = "SOTINVH1.ORDR_YYYYPP_UPDATED"
        sql_Data = "" _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & LYP & "' THEN " & sqlX & " ELSE 0 END) RETURNED_LYTM" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & LYP01 & "' AND " & sqlYP & " <= '" & LYP & "' THEN " & sqlX & " ELSE 0 END) RETURNED_LYTD" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' THEN " & sqlX & " ELSE 0 END) RETURNED_TYTM" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & TYP01 & "' AND " & sqlYP & " <= '" & TYP & "' THEN " & sqlX & " ELSE 0 END) RETURNED_TYTD" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " >= '" & LYP01 & "' AND " & sqlYP & " <= '" & LYP12 & "' THEN " & sqlX & " ELSE 0 END) RETURNED_LYTOT" & vbCrLf _
            & ", SUM (CASE WHEN " & sqlYP & " = '" & TYP & "' AND SOTINVH1.INV_DATE = '" & Format(Now, "dd-MMM-yyyy") & "' THEN " & sqlX & " ELSE 0 END) RETURNED_TODAY" & vbCrLf

        sql_Cols = "" _
            & ",RETURNED_LYTM,RETURNED_LYTD,RETURNED_TYTM,RETURNED_TYTD,RETURNED_LYTOT,RETURNED_TODAY"

        sql_filter = "" _
            & " and " & sqlYP & " BETWEEN '" & LYP01 & "' and '" & TYP & "'" _
            & " and SOTINVH2.INV_TYPE = 'C'"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTINVH1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

 

        ASCMAIN1.Progress("Budgets")
        MyBase.Get_SQL("B")

        If ASCMAIN1.CLIENT = "AHA" Then

            sql_Data = "" _
                & ", SUM ((CASE WHEN SATBUDG1.OPS_YYYYPP = '" & TYP & "' THEN SATBUDG1.BUDGET ELSE 0 END)" & FX & ") BUDGET_TYTM" & vbCrLf _
                & ", SUM ((CASE WHEN SATBUDG1.OPS_YYYYPP >= '" & TYP01 & "' AND SATBUDG1.OPS_YYYYPP <= '" & TYP & "' THEN SATBUDG1.BUDGET ELSE 0 END)" & FX & ") BUDGET_TYTD" & vbCrLf _
                & ", SUM ((CASE WHEN SATBUDG1.OPS_YYYYPP >= '" & TYP01 & "' AND SATBUDG1.OPS_YYYYPP <= '" & TYP12 & "' THEN SATBUDG1.BUDGET ELSE 0 END)" & FX & ") BUDGET_TYTOT" & vbCrLf

            sql_Cols = "" _
                & ",BUDGET_TYTM,BUDGET_TYTD,BUDGET_TYTOT"

            sql_filter = "" _
                & " and SATBUDG1.OPS_YYYYPP >= '" & TYP01 & "' AND SATBUDG1.OPS_YYYYPP <= '" & TYP12 & "'"

            sql = "Select " & sql_SELECT_cols & vbCrLf _
                & "" & vbCrLf & sql_Data _
                & " from SATBUDG1" & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

        Else
            sqlX = ""
            sqlX12 = ""
            For I As Integer = 1 To 12
                If I <= Val(Mid(TYP, 5, 2)) Then sqlX &= "+NVL(SATBUDW1.WB_P" & Format(I, "00") & ",0)" & FX
                sqlX12 &= "+NVL(SATBUDW1.WB_P" & Format(I, "00") & ",0)" & FX
            Next

            sql_Data = "" _
                & ", SUM (NVL(SATBUDW1.WB_P" & Format(Val(Mid(TYP, 5, 2)), "00") & ",0)" & FX & ") BUDGET_TYTM" & vbCrLf _
                & ", SUM (" & Mid(sqlX, 2) & ") BUDGET_TYTD" & vbCrLf _
                & ", SUM (" & Mid(sqlX12, 2) & ") BUDGET_TYTOT" & vbCrLf

            sql_Cols = "" _
                & ",BUDGET_TYTM,BUDGET_TYTD,BUDGET_TYTOT"

            sql_filter = "" _
                & " and SATBUDW1.OPS_YYYY = '" & Mid(TYP, 1, 4) & "'"

            sql = "Select " & sql_SELECT_cols & vbCrLf _
                & "" & vbCrLf & sql_Data _
                & " from " & SATBUDW1 & " SATBUDW1" & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

        End If

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()



        'ASCMAIN1.Progress("Budgets for Returns")
        'MyBase.Get_SQL("B")

        'sqlX = ""
        'sqlX12 = ""
        'For I As Integer = 1 To 12
        '    If I <= Val(Mid(TYP, 5, 2)) Then sqlX &= "+NVL(SATBUDW1.WB_P" & Format(I, "00") & ",0)" & FX
        '    sqlX12 &= "+NVL(SATBUDW1.WB_P" & Format(I, "00") & ",0)" & FX
        'Next

        'sql_Data = "" _
        '    & ", SUM (NVL(SATBUDW1.WB_P" & Format(Val(Mid(TYP, 5, 2)), "00") & ",0)" & FX & ") BUDGET_RTN_TYTM" & vbCrLf _
        '    & ", SUM (" & Mid(sqlX, 2) & ") BUDGET_RTN_TYTD" & vbCrLf _
        '    & ", SUM (" & Mid(sqlX12, 2) & ") BUDGET_RTN_TOT" & vbCrLf

        'sql_Cols = "" _
        '    & ",BUDGET_RTN_TYTM,BUDGET_RTN_TYTD,BUDGET_RTN_TYTOT"

        'sql_filter = "" _
        '    & " and SATBUDW1.OPS_YYYY = '" & Mid(TYP, 1, 4) & "'"

        'sql = "Select " & sql_SELECT_cols & vbCrLf _
        '    & "" & vbCrLf & sql_Data _
        '    & " from " & SATBUDW1 & " SATBUDW1" & sql_TABLE_NAMEs & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        '    & " group by " & sql_GROUP_BY_cols

        'ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        '    & "(" & G1thru9 & COLUMN_NAMEs_appended _
        '    & sql_Cols & ")" & vbCrLf _
        '    & "(" & sql & ")"
        'ASCDATA1.ExecuteSQL()



        ' Eliminate 0s

        Dim sqlz As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            sqlz &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlz))
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "Y", "N"))
        Dim pp = Val(Mid(ASCMAIN1.CYP, 5, 2))
        CR_params.Add("OTS1", Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0)), 10, 3))
        CR_params.Add("OTS2", Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)), 10, 3))
        CR_params.Add("OTS3", Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 2)), 10, 3) & "+")
        CR_params.Add("REPORT", Absx1.optFor("REPORT").Value)
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            EMsg &= TAC.TACMAIN1.Check_Permissions(Me) ' for FS
        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        With dst.Tables("ASTSRPT1")
            .Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
            .Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
            .Columns("PROJECTED_TYTD").Expression = "ISNULL(SHIPPED_TYTD,0)+ISNULL(OTS_M01,0)"
            .Columns("BUDGET_TYTOGO").Expression = "ISNULL(BUDGET_TYTOT,0)-ISNULL(BUDGET_TYTD,0)"
            .Columns("BUDGET_RTN_TYTOGO").Expression = "ISNULL(BUDGET_RTN_TYTOT,0)-ISNULL(BUDGET_RTN_TYTD,0)"
            .Columns("SHIPPED_LYTOGO").Expression = "ISNULL(SHIPPED_LYTOT,0)-ISNULL(SHIPPED_LYTD,0)"
            .Columns("RETURNED_LYTOGO").Expression = "ISNULL(RETURNED_LYTOT,0)-ISNULL(RETURNED_LYTD,0)"
            .Columns("NET_SHP_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)-ISNULL(RETURNED_TYTM,0)"
            .Columns("NET_SHP_TYTD").Expression = "ISNULL(SHIPPED_TYTD,0)-ISNULL(RETURNED_TYTD,0)"
            .Columns("NET_BUD_TYTM").Expression = "ISNULL(BUDGET_TYTM,0)-ISNULL(BUDGET_RTN_TYTM,0)"
            .Columns("NET_BUD_TYTD").Expression = "ISNULL(BUDGET_TYTD,0)-ISNULL(BUDGET_RTN_TYTD,0)"
            .Columns("NET_BUD_TYTOT").Expression = "ISNULL(BUDGET_TYTOT,0)-ISNULL(BUDGET_RTN_TYTOT,0)"
            .Columns("NET_BUD_TYTOGO").Expression = "ISNULL(NET_BUD_TYTOT,0)-ISNULL(NET_BUD_TYTD,0)"
            .Columns("NET_SHP_LYTM").Expression = "ISNULL(SHIPPED_LYTM,0)-ISNULL(RETURNED_LYTM,0)"
            .Columns("NET_SHP_LYTD").Expression = "ISNULL(SHIPPED_LYTD,0)-ISNULL(RETURNED_LYTD,0)"
            .Columns("NET_SHP_LYTOT").Expression = "ISNULL(SHIPPED_LYTOT,0)-ISNULL(RETURNED_LYTOT,0)"
            .Columns("NET_SHP_LYTOGO").Expression = "ISNULL(NET_SHP_LYTOT,0)-ISNULL(NET_SHP_LYTD,0)"
            '.Columns("BOOK_TYTOGO").Expression = "ISNULL(BUDGET_TYTOT,0)-ISNULL(BUDGET_TYTD,0)"
        End With

        If chkFLASH.Checked Then Create_Pivot()
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("SREP_CODE")
                ' rowASTDSQLA.Item("CODE_VALUES") = TAC.TACMAIN1.SREP_CODE
                rowASTDSQLA.Item("CODE_VALUES") = Join(TAC.TACMAIN1.SREP_CODEs.ToArray, ",")
            End If
        End If
    End Sub

    Sub Create_Pivot()

        Dim SQLW As String = ""

        ASCMAIN1.Progress("Now Creating Workbook")
        'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
        '    ASCMAIN1.Folders("SharedRoot") = "C:\Share\INT\"
        'End If
        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"

        If ASCMAIN1.Running_in_VS And Format(Now, "MM/dd/yyyy") = "01/13/2025" Then
            Stop
            FILENAME = "C:\Share\INT\Templates\" & Me.Name & ".xlsx"
        End If

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim SheetName As String = "Data"
        ws = wb.Worksheets(SheetName)

        Dim DataTable As DataTable = dst.Tables("ASTSRPT1")

        Dim iRx As Integer = 1

        Dim r As Integer = iRx
        For Each row As DataRow In DataTable.Select("")
            r += 1
            ws.Range(Excel_Cell(r, 1) & ":" & Excel_Cell(r, DataTable.Columns.Count)).Value2 = row.ItemArray
        Next
 
        ASCMAIN1.Progress("-", "Pivot")
        wb.Names.Add("PivotBase", "=" & SheetName & "!" & Excel_Cell(iRx, 1, 3) & ":" & Excel_Cell(iRx + DataTable.Rows.Count, DataTable.Columns.Count, 3))
        'excel.Run("ResetData")

        'refresh the pivotcache
        'ws.PivotTables("PivotTable1").PivotCache.Refresh()

        ' the line below was disabled by wjz on 02/08 as part of disabling all refreshes -but not sure if this one was a problem - this one might have worked
        '   wb.Sheets("YTD").PivotTables(1).PivotCache.Refresh()

        'Dim pt As Microsoft.Office.Interop.Excel.PivotTable
        'pt = wb.Sheets("YTD").PivotTables("PivotTable1")
        'pt.PivotCache.Refresh()

        'Marshal.ReleaseComObject(pt)

        'ws.PivotTables("PivotTable1").RefreshTable()
        'ws.PivotTables("PivotTable1").Update()

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Daily_Sales_Flash"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"

                Dim try_again As Boolean = False
                If ASCMAIN1.JOB_STREAM_CODE <> "" Then
                    If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Work") & XLS_FILENAME) Then
                        try_again = True
                    End If
                End If

                If Not try_again Then
                    Dim objOpt As Object = Nothing ' Missing.Value
                    wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                              , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                    wb.Close(False, objOpt, objOpt)

                    success = True
                End If

            Catch ex As Exception
                ' Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing
        ' pt = Nothing

        If ASCMAIN1.JOB_STREAM_CODE = "" Then
            ' ReleaseCOMObject(pt)
            ReleaseCOMObject(xlDestRange)
            ReleaseCOMObject(xlSourceRange)
            ReleaseCOMObject(ws)
            ReleaseCOMObject(wb)
            ReleaseCOMObject(excel)
        End If
        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME, (ASCMAIN1.JOB_STREAM_CODE = ""))
        If ASCMAIN1.JOB_STREAM_CODE <> "" Then
            grdASTSRPT1.DataSource = dst.Tables("ASTSRPT1")
            Dim DATA_GRID_FILENAME As String = Excel_Export_to_SSG(grdASTSRPT1)
            Add_Document_to_ASTSPRF1(DATA_GRID_FILENAME, (ASCMAIN1.JOB_STREAM_CODE = ""))
        End If
        ASCMAIN1.Progress("")
    End Sub

    Private Sub btnPivot_Click(sender As Object, e As EventArgs)
        Create_Pivot()
    End Sub

    Private Sub chkFLASH_CheckedChanged(sender As Object, e As EventArgs) Handles chkFLASH.CheckedChanged
        grpFlash.Visible = chkFLASH.Checked
    End Sub
     
End Class