Public Class SPRAOPN1

    Dim md(12) As String
    Dim SPTACOMC As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        Dim z As String = String.Empty
        Dim zz As String = String.Empty

        Dim endPeriod As String = cmbRYP.SelectedRow.Cells(0).Value
        Dim startPeriod As String = ASCMAIN1.Period_Calc(endPeriod, -11)

        ASCMAIN1.sql = "Select SPTACOMC.*, NVL(SPTACOMC_PRE.OPS_YYYYPP_PAID, SPTACOMC.OPS_YYYYPP) OPS_YYYYPP_OPEN" _
            & " from SPTACOMC, SPTACOMC SPTACOMC_PRE where SPTACOMC_PRE.ACC_CTL_NO (+) = SPTACOMC.ACC_CTL_NO_ORIG"
        SPTACOMC = ASCMAIN1.Temp_Table

        'Dim startDate As String = startPeriod.Substring(4, 2) & "/01/" & startPeriod.Substring(0, 4)
        'Dim endDate As String = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(endPeriod.Substring(4, 2) & "/01/" & endPeriod.Substring(0, 4)))).ToString("MM/dd/yyyy")

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
            sql &= ", SUM (CASE WHEN OPS_YYYYPP_OPEN <= '" & iYP & "' AND (OPS_YYYYPP_PAID IS NULL OR OPS_YYYYPP_PAID > '" & iYP & "') THEN AMT_COMM ELSE 0 END) V" & (i + 1).ToString("00") & vbCrLf
        Next

        sql &= " from " & SPTACOMC & " SPTACOMC" & vbCrLf
        sql &= " where SPTACOMC.OPS_YYYYPP <= '" & endPeriod & "'" & vbCrLf
        sql &= sql_JOIN
        sql &= sql_WHERE
        sql &= " group by " & sql_GROUP_BY_cols & ", SPTACOMC.OPS_YYYYPP" & vbCrLf

         sql = "Insert Into " & ASTSRPT1 & " " & Sql
        ASCDATA1.ExecuteSQL(Sql)
   
        ASCMAIN1.Progress("", "")
    End Sub

    Public Overrides Sub Print_Report()
        For i As Integer = 1 To 12
            CR_params.Add("MD" & i.ToString("00"), md(i))
        Next i

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "Y", "N"))
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
End Class