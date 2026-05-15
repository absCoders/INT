Public Class ASFWCAL1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "Select 0 LINE_NO, 'X' MONTH, '99' WEEK_NO"
            For i As Integer = 1 To 5   ' 5 years
                For j As Integer = 1 To 7  ' 7 days in a week
                    ASCMAIN1.sql &= ", 'XX' DATE_" & Format$(i, "0") & Format$(j, "0")
                Next j
            Next i
            ASCMAIN1.sql &= " from GLTPARM3"

            Create_TDA(.Tables.Add, "ASTWCAL1", "**", 0, False, "", 1)
        End With

        grdASTWCAL1.DataSource = dst.Tables("ASTWCAL1")

        Dim CW As Integer = 25
        Dim WEEKDAYS As String = "SMTWRFS"

        With grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1")

            .Groups("WEEK").Header.Appearance.TextHAlign = HAlign.Center
            .Groups("WEEK").Header.Caption = ""
            .Groups("WEEK").Width = CW * 2
            .Groups("WEEK").Header.Appearance.BackColor = Drawing.Color.DodgerBlue
            .Groups("WEEK").Header.Appearance.BackColor2 = Drawing.Color.Empty
            .Groups("WEEK").Header.Appearance.BackGradientStyle = GradientStyle.None
            .Groups("WEEK").Header.Appearance.ForeColor = Drawing.Color.White

            .Columns("LINE_NO").Hidden = True

            .Columns("MONTH").Width = CW
            .Columns("MONTH").Header.Caption = "M"
            .Columns("MONTH").Header.Appearance.TextHAlign = HAlign.Center
            .Columns("MONTH").CellAppearance.TextHAlign = HAlign.Center
            .Columns("MONTH").Header.Appearance.ForeColor = Drawing.Color.DodgerBlue
            .Columns("MONTH").CellAppearance.ForeColor = Drawing.Color.DarkBlue

            .Columns("WEEK_NO").Width = CW
            .Columns("WEEK_NO").Header.Caption = "W"
            .Columns("WEEK_NO").Header.Appearance.TextHAlign = HAlign.Center
            .Columns("WEEK_NO").CellAppearance.TextHAlign = HAlign.Center

            .Columns("WEEK_NO").Header.Appearance.ForeColor = Drawing.Color.DodgerBlue
            .Columns("WEEK_NO").CellAppearance.BackColor = Drawing.Color.AliceBlue

            .Columns("WEEK_NO").CellAppearance.ForeColor = Drawing.Color.DarkBlue
            '.Columns("WEEK_NO").CellAppearance.BackColor = Drawing.Color.DodgerBlue
            '.Columns("WEEK_NO").CellAppearance.ForeColor = Drawing.Color.White

            For i As Integer = 1 To 5   ' 5 years
                Dim G As String = "Y" & Format(i, "0")
                .Groups(G).Header.Appearance.TextHAlign = HAlign.Center
                .Groups(G).Width = CW * 7.5
                .Groups(G).Header.Appearance.BackColor = Drawing.Color.DodgerBlue
                .Groups(G).Header.Appearance.BackColor2 = Drawing.Color.Empty
                .Groups(G).Header.Appearance.BackGradientStyle = GradientStyle.None
                .Groups(G).Header.Appearance.ForeColor = Drawing.Color.White

                For j As Integer = 1 To 7  ' 7 days in a week
                    Dim C As String = "DATE_" & Format(i, "0") & Format(j, "0")
                    .Columns(C).Width = CW
                    .Columns(C).Header.Caption = Mid(WEEKDAYS, j, 1)
                    .Columns(C).Header.Appearance.TextHAlign = HAlign.Center
                    .Columns(C).CellAppearance.TextHAlign = HAlign.Center
                    If j = 7 Then
                        .Columns(C).CellAppearance.BackColor = Drawing.Color.Yellow
                        .Columns(C).CellAppearance.BackColor2 = Drawing.Color.White
                        .Columns(C).CellAppearance.BackGradientStyle = GradientStyle.GlassRight20
                    End If
                Next j
            Next i
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Show Calendar"
                Dim YYYY As String = Absx1.txtFor("YYYY").Text
                If Len(YYYY) <> 4 Or Not IsNumeric(YYYY) Then
                    EMsg = EMsg & vbCr & "Invalid Year (" & YYYY & ")"
                Else
                    If Val(YYYY) > Val(Mid(ASCMAIN1.CYP, 1, 4)) + 20 Or Val(YYYY) < Val(Mid(ASCMAIN1.CYP, 1, 4)) - 20 Then
                        EMsg = EMsg & vbCr & "Invalid Year (" & YYYY & ")"
                    End If
                End If


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Show Calendar"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()

            Case "Excel"
                Call Export_to_Excel(grdASTWCAL1)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Show Calendar").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdASTWCAL1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        Absx1.txtFor("YYYY").Text = Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) - 2, "0000")
        dst.Tables("ASTWCAL1").Rows.Clear()
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        Dim Y1 As String = HFs("YYYY")
        Dim Y5 As String = Format(Val(HFs("YYYY")) + 4, "0000")

        Dim Sql As String = ""

        Sql = "Select SUBSTR(YYYYWW,5,2) WW, SUBSTR(YYYYMM,5,2) MM" _
            & ", REL_WEEK, MAX(MAX_WEEK) MAX_WEEK"

        For Y As Integer = 1 To 5
            Sql = Sql & " , MIN(DECODE (SUBSTR(YYYYWW,1,4),'" & Format$(Val(Y1) + Y - 1, "0000") & "',WEEK_END_DATE,NULL)) Y" & Format$(Y, "0")
        Next Y

        Sql &= " From GLTPARM3" _
            & " where YYYYWW >= '" & Y1 & "01'" _
            & "   and YYYYWW <= '" & Y5 & "53'" _
            & " GROUP BY SUBSTR(YYYYWW,5,2), SUBSTR(YYYYMM,5,2), REL_WEEK" _
            & " ORDER BY SUBSTR(YYYYWW,5,2), SUBSTR(YYYYMM,5,2), REL_WEEK"

        Dim LINE_NO As Integer = 0

        For Each row As DataRow In ASCDATA1.GetDataTable(Sql).Select("", "WW,MM,REL_WEEK")
            Dim REL_WEEK As Integer = Val(row.Item("REL_WEEK") & "")
            Dim MAX_WEEK As Integer = Val(row.Item("MAX_WEEK") & "")
            Dim MM As String = row.Item("MM")
            Dim WW As String = row.Item("WW")
            Dim MONTH_X As String = Mid$(" " & Format(DateValue(MM & "/01/2000"), "MMM") & " ", REL_WEEK, 1)

            Dim rowASTWCAL1 As DataRow = dst.Tables("ASTWCAL1").NewRow
            LINE_NO += 1
            rowASTWCAL1.Item("LINE_NO") = LINE_NO
            rowASTWCAL1.Item("MONTH") = MONTH_X
            rowASTWCAL1.Item("WEEK_NO") = WW

            For Y As Integer = 1 To 5
                If row.Item("Y" & Format$(Y, "0")) & "" <> "" Then
                    Dim dt As Date = row.Item("Y" & Format$(Y, "0"))
                    For d As Integer = 1 To 7
                        rowASTWCAL1.Item("DATE_" & Format(Y, "0") & Format(d, "0")) = Format(dt.AddDays(-7 + d), "dd")
                    Next d
                End If
            Next Y
            dst.Tables("ASTWCAL1").Rows.Add(rowASTWCAL1)

            If REL_WEEK = MAX_WEEK Then
                rowASTWCAL1 = dst.Tables("ASTWCAL1").NewRow
                LINE_NO += 1
                rowASTWCAL1.Item("LINE_NO") = LINE_NO
                dst.Tables("ASTWCAL1").Rows.Add(rowASTWCAL1)
            End If
            'Stop

        Next

        For i As Integer = 1 To 5   ' 5 years
            Dim G As String = "Y" & Format(i, "0")
            grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1").Groups(G).Header.Caption = Format(Y1 + i - 1, "0000")
        Next i


        With grdASTWCAL1.DisplayLayout.Bands("ASTWCAL1").SortedColumns
            .Clear()
            .Add("LINE_NO", False)
        End With

    End Sub

    Sub Update_Record()
        Call BeginTrans()

        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "YYYY"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Show Calendar", e)
                End If
        End Select
    End Sub
#End Region

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim YYYY As Integer = Val(HFs("YYYY"))
        CR_params.Add("Y1", Format(YYYY + 0, "0000"))
        CR_params.Add("Y2", Format(YYYY + 1, "0000"))
        CR_params.Add("Y3", Format(YYYY + 2, "0000"))
        CR_params.Add("Y4", Format(YYYY + 3, "0000"))
        CR_params.Add("Y5", Format(YYYY + 4, "0000"))
        Generate_Report("ASRWCAL1", "Weekly Retail Calendar")

        Call Print_Report_End()
    End Sub

    Private Sub grdASTWCAL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTWCAL1.InitializeRow
        If e.Row.Cells("WEEK_NO").Text = "" Then
            e.Row.Appearance.BackColor = Drawing.Color.AliceBlue
            e.Row.Height = 10
        End If
    End Sub
End Class