Public Class SPRCSUM1

    Private md(12) As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        Dim z As String = String.Empty
        Dim zz As String = String.Empty

        Dim endPeriod As String = cmbRYP.SelectedRow.Cells(0).Value
        Dim startPeriod As String = ASCMAIN1.Period_Calc(endPeriod, -11)

        Dim startDate As String = startPeriod.Substring(4, 2) & "/01/" & startPeriod.Substring(0, 4)
        Dim endDate As String = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(endPeriod.Substring(4, 2) & "/01/" & endPeriod.Substring(0, 4)))).ToString("MM/dd/yyyy")

        MyBase.Get_SQL("*")

        ReDim md(12)
        Dim captionPeriod As String = startPeriod
        For i As Integer = 1 To 12
            z = captionPeriod.Substring(4, 2) & "/01/" & captionPeriod.Substring(0, 4)
            md(i) = CDate(z).ToString("MMM yy")
            md(i) = md(i).Replace(" ", "-")
            captionPeriod = ASCMAIN1.Period_Calc(captionPeriod, 1)
        Next i

        sql = "Select " & sql_SELECT_cols & vbCr

        For i As Integer = 0 To 11
            Dim iYP As String = ASCMAIN1.Period_Calc(startPeriod, i)
            Select Case optOpenPaid.Value
                Case "O"
                    sql &= ", Sum (DECODE(SPTCOOP1.OPS_YYYYPP, '" & iYP & "', (NVL(SPTCOOP1.OPEN_AMT, 0) * NVL(SPTCOOP3.DIST_AMT, 0) / (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000))), 0)) V" & (i + 1).ToString("00") & vbCrLf

                Case "P"
                    sql &= ", Sum (DECODE(SPTCOOP1.OPS_YYYYPP, '" & iYP & "', NVL(SPTCOOP1.PAID_AMT, 0) * NVL(SPTCOOP3.DIST_AMT, 0) / (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000)), 0)) V" & (i + 1).ToString("00") & vbCrLf

                Case "A"
                    sql &= ", Sum (DECODE(SPTCOOP1.OPS_YYYYPP, '" & iYP & "', NVL(SPTCOOP1.OPEN_AMT, 0) * NVL(SPTCOOP3.DIST_AMT, 0) / (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000)) "
                    sql &= "+ NVL(SPTCOOP1.PAID_AMT, 0) * NVL(SPTCOOP3.DIST_AMT, 0) / (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000)), 0)) V" & (i + 1).ToString("00") & vbCrLf
            End Select
        Next
        sql &= ", 0 PTD, 0 TOT, 0 TOT_PCT, 0 PTD_PCT" & vbCr

        sql &= " from SPTCOOP1, SPTCOOP3" & sql_TABLE_NAMEs.Replace(",SPTCOOP3", "") & vbCr
        sql &= " where SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO " & vbCr
        sql &= "   and SPTCOOP1.OPS_YYYYPP >= '" & startPeriod & "'" & vbCr
        sql &= "   and SPTCOOP1.OPS_YYYYPP <= '" & endPeriod & "'" & vbCr
        sql &= "   and (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000)) > 0" & vbCr
        sql &= sql_JOIN
        sql &= sql_WHERE
        sql &= " group by " & sql_GROUP_BY_cols & ", SPTCOOP1.OPS_YYYYPP" & vbCr

        sql = "Insert Into " & ASTSRPT1 & " " & sql
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress("Now Calculating Totals", "")

        '*********************************************************************
        sql = String.Empty
        For i As Integer = 1 To 12
            sql &= "+ NVL(V" & i.ToString("00") & ", 0) "
        Next i

        sql = sql.Substring(2)
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT = " & sql)

        '*********************************************************************
        sql = String.Empty
        For i As Integer = 1 To numMonths.Value
            sql &= "+ NVL(V" & i.ToString("00") & ", 0) "
        Next i
        sql = sql.Substring(2)
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set PTD = " & sql)

        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT_PCT = 0, PTD_PCT = 0")
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT_PCT = 100 WHERE TOT > 0")
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set PTD_PCT = 100 WHERE PTD > 0")

        If COLUMN_NAMEs.Count > 1 Then

            '*********************************************************************
            sql = "Select "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " G" & i & "," & vbCr
            Next
            sql &= " SUM("
            For i As Integer = 1 To 12
                sql &= " V" & i.ToString("00") & "+"
            Next

            sql = sql.Substring(0, sql.Length - 1)
            sql &= ") TOTAL"

            sql &= " FROM " & ASTSRPT1
            sql &= " GROUP BY "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " G" & i & ","
            Next

            sql = sql.Substring(0, sql.Length - 1)
            Dim totTable As String = ASCMAIN1.Temp_Table(sql)


            sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & totTable & ";"
            sql &= " BEGIN FOR R1 IN C1 LOOP "
            sql &= "   UPDATE " & ASTSRPT1 & " SET TOT_PCT = (TOT / R1.TOTAL) * 100 WHERE R1.TOTAL <> 0"
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " AND G" & i & " = R1.G" & i
            Next
            sql &= ";"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

            '*********************************************************************
            sql = "Select "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " G" & i & "," & vbCr
            Next
            sql &= " SUM("
            For i As Integer = 1 To numMonths.Value
                sql &= " V" & i.ToString("00") & "+"
            Next

            sql = sql.Substring(0, sql.Length - 1)
            sql &= ") TOTAL"

            sql &= " FROM " & ASTSRPT1
            sql &= " GROUP BY "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " G" & i & ","
            Next

            sql = sql.Substring(0, sql.Length - 1)
            totTable = ASCMAIN1.Temp_Table(sql)

            sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & totTable & ";"
            sql &= " BEGIN FOR R1 IN C1 LOOP "
            sql &= "   UPDATE " & ASTSRPT1 & " SET PTD_PCT = (TOT / R1.TOTAL) * 100 WHERE R1.TOTAL <> 0"
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                sql &= " AND G" & i & " = R1.G" & i
            Next
            sql &= ";"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)
        End If

        ASCMAIN1.Progress("", "")
    End Sub

    Public Overrides Sub Print_Report()
        For i As Integer = 1 To 12
            CR_params.Add("MD" & i.ToString("00"), md(i))
        Next i

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "Y", "N"))
        CR_params.Add("MD1", md(1))
        CR_params.Add("MD2", md(numMonths.Value))
        CR_params.Add("T1", "1")
        CR_params.Add("T2", CStr(numMonths.Value))

        If numMonths.Value = 12 Then
            CR_params.Add("HIDE", "0")
        Else
            CR_params.Add("HIDE", "1")
        End If

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

End Class