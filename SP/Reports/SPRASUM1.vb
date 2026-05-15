Public Class SPRASUM1

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

        Sql = "Select " & sql_SELECT_cols & vbCr

        For i As Integer = 0 To 11
            Dim iYP As String = ASCMAIN1.Period_Calc(startPeriod, i)
            Select Case optOpenPaid.Value
                Case "A"
                    sql &= ", Sum (DECODE(SPTACOMC.OPS_YYYYPP, '" & iYP & "', NVL(SPTACOMC.AMT_COMM, 0))) V" & (i + 1).ToString("00") & vbCrLf

                Case "P"
                    sql &= ", Sum (DECODE(SPTACOMC.OPS_YYYYPP, '" & iYP & "', NVL(SPTACOMC.AMT_COMM_PAID, 0))) V" & (i + 1).ToString("00") & vbCrLf

                Case "O"
                    sql &= ", Sum (DECODE(SPTACOMC.OPS_YYYYPP, '" & iYP & "', NVL(SPTACOMC.AMT_COMM, 0) - NVL(SPTACOMC.AMT_COMM_PAID, 0))) V" & (i + 1).ToString("00") & vbCrLf
            End Select
        Next
        sql &= ", 0 PTD, 0 TOT, 0 TOT_PCT, 0 PTD_PCT" & vbCrLf

        sql &= " from SPTACOMC" & vbCrLf
        sql &= " where SPTACOMC.OPS_YYYYPP >= '" & startPeriod & "'" & vbCrLf
        sql &= "   and SPTACOMC.OPS_YYYYPP <= '" & endPeriod & "'" & vbCrLf
        sql &= sql_JOIN
        sql &= sql_WHERE
        sql &= " group by " & sql_GROUP_BY_cols & ", SPTACOMC.OPS_YYYYPP" & vbCrLf

        '        sql &= "   and (NVL(SPTCOOP1.OTHER_COST, 0) + (NVL(SPTCOOP1.QTY, 0) * NVL(SPTCOOP1.VEHICLE_CPM, 0) / 1000)) > 0" & vbCrLf

        sql = "Insert Into " & ASTSRPT1 & " " & Sql
        ASCDATA1.ExecuteSQL(Sql)

        ASCMAIN1.Progress("Now Calculating Totals", "")

        '*********************************************************************
        Sql = String.Empty
        For i As Integer = 1 To 12
            Sql &= "+ NVL(V" & i.ToString("00") & ", 0) "
        Next i

        Sql = Sql.Substring(2)
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT = " & Sql)

        '*********************************************************************
        Sql = String.Empty
        For i As Integer = 1 To numMonths.Value
            Sql &= "+ NVL(V" & i.ToString("00") & ", 0) "
        Next i
        Sql = Sql.Substring(2)
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set PTD = " & Sql)

        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT_PCT = 0, PTD_PCT = 0")
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set TOT_PCT = 100 WHERE TOT > 0")
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " set PTD_PCT = 100 WHERE PTD > 0")

        If COLUMN_NAMEs.Count > 1 Then

            '*********************************************************************
            Sql = "Select "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " G" & i & "," & vbCr
            Next
            Sql &= " SUM("
            For i As Integer = 1 To 12
                Sql &= " V" & i.ToString("00") & "+"
            Next

            Sql = Sql.Substring(0, Sql.Length - 1)
            Sql &= ") TOTAL"

            Sql &= " FROM " & ASTSRPT1
            Sql &= " GROUP BY "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " G" & i & ","
            Next

            Sql = Sql.Substring(0, Sql.Length - 1)
            Dim totTable As String = ASCMAIN1.Temp_Table(Sql)


            Sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & totTable & ";"
            Sql &= " BEGIN FOR R1 IN C1 LOOP "
            Sql &= "   UPDATE " & ASTSRPT1 & " SET TOT_PCT = (TOT / R1.TOTAL) * 100 WHERE R1.TOTAL <> 0"
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " AND G" & i & " = R1.G" & i
            Next
            Sql &= ";"
            Sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(Sql)

            '*********************************************************************
            Sql = "Select "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " G" & i & "," & vbCr
            Next
            Sql &= " SUM("
            For i As Integer = 1 To numMonths.Value
                Sql &= " V" & i.ToString("00") & "+"
            Next

            Sql = Sql.Substring(0, Sql.Length - 1)
            Sql &= ") TOTAL"

            Sql &= " FROM " & ASTSRPT1
            Sql &= " GROUP BY "
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " G" & i & ","
            Next

            Sql = Sql.Substring(0, Sql.Length - 1)
            totTable = ASCMAIN1.Temp_Table(Sql)

            Sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & totTable & ";"
            Sql &= " BEGIN FOR R1 IN C1 LOOP "
            Sql &= "   UPDATE " & ASTSRPT1 & " SET PTD_PCT = (TOT / R1.TOTAL) * 100 WHERE R1.TOTAL <> 0"
            For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                Sql &= " AND G" & i & " = R1.G" & i
            Next
            Sql &= ";"
            Sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(Sql)
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