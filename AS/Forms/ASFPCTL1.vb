Public Class ASFPCTL1

    ' https://nrf.com/resources/4-5-4-calendar

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ASTPCTL1", "*")
            Create_TDA(.Tables.Add, "GLTPARM2", "*", 0)
            Create_TDA(.Tables.Add, "GLTPARM3", "*", 0)

            ASCMAIN1.sql = "SELECT SUBSTR(OPS_YYYYPP,1,4) YEAR"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", MAX(DECODE(SUBSTR(OPS_YYYYPP,5,2),'" & Format(M, "00") & "',PRD_END_DATE,NULL)) END_" & Format(M, "00")
            Next
            ASCMAIN1.sql &= " FROM GLTPARM2 GROUP BY SUBSTR(OPS_YYYYPP,1,4)"
            Create_TDA(.Tables.Add, "GLTPARM2X", "**", 0, False, "", 1)
            For M As Integer = 1 To 12
                dst.Tables("GLTPARM2X").Columns("END_" & Format(M, "00")).DataType = GetType(System.DateTime)
            Next
            'With .Tables.Add("GLTPARM2X")
            '    .Columns.Add("YEAR")
            '    For I As Integer = 1 To 12
            '        .Columns.Add("END_" & Format(I, "00"), GetType(System.DateTime))
            '    Next
            'End With

            ASCMAIN1.sql = "SELECT SUBSTR(YYYYWW,1,4) YEAR"
            For W As Integer = 1 To 53
                ASCMAIN1.sql &= ", MAX(DECODE(SUBSTR(YYYYWW,5,2),'" & Format(W, "00") & "',WEEK_END_DATE,NULL)) END_" & Format(W, "00")
            Next
            ASCMAIN1.sql &= " FROM GLTPARM3 GROUP BY SUBSTR(YYYYWW,1,4)"
            Create_TDA(.Tables.Add, "GLTPARM3X", "**", 0, False, "", 1)
            For M As Integer = 1 To 53
                dst.Tables("GLTPARM3X").Columns("END_" & Format(M, "00")).DataType = GetType(System.DateTime)
            Next

            'With .Tables.Add("GLTPARM3X")
            '    .Columns.Add("YEAR")
            '    For I As Integer = 1 To 53
            '        .Columns.Add("END_" & Format(I, "00"), GetType(System.DateTime))
            '    Next
            'End With

            With .Tables.Add("ASTPCTLP")
                .Columns.Add("PERIOD_END_DAY")
                .Columns.Add("PERIOD_END_DAY_DESC")
                .Rows.Add(New String() {"0", "Calendar Day"})
                .Rows.Add(New String() {"1", "Sunday"})
                .Rows.Add(New String() {"2", "Monday"})
                .Rows.Add(New String() {"3", "Tuesday"})
                .Rows.Add(New String() {"4", "Wednesday"})
                .Rows.Add(New String() {"5", "Thursday"})
                .Rows.Add(New String() {"6", "Friday"})
                .Rows.Add(New String() {"7", "Saturday"})
            End With

            With .Tables.Add("ASTPCTLW")
                .Columns.Add("WEEK_END_DAY")
                .Columns.Add("WEEK_END_DAY_DESC")
                .Rows.Add(New String() {"1", "Sunday"})
                .Rows.Add(New String() {"2", "Monday"})
                .Rows.Add(New String() {"3", "Tuesday"})
                .Rows.Add(New String() {"4", "Wednesday"})
                .Rows.Add(New String() {"5", "Thursday"})
                .Rows.Add(New String() {"6", "Friday"})
                .Rows.Add(New String() {"7", "Saturday"})
            End With

        End With

        grdGLTPARM2X.DataSource = dst.Tables("GLTPARM2X")
        grdGLTPARM3X.DataSource = dst.Tables("GLTPARM3X")


        '            Case "PERIOD_END_DAY"
        'VL.Add("0", "Calendar")
        'VL.Add("1", "Sunday")
        'VL.Add("2", "Monday")
        'VL.Add("3", "Tuesday")
        'VL.Add("4", "Wednesday")
        'VL.Add("5", "Thursday")
        'VL.Add("6", "Friday")
        'VL.Add("7", "Saturday")

        '    Case "WEEK_END_DAY"
        'VL.Add("1", "Sunday")
        'VL.Add("2", "Monday")
        'VL.Add("3", "Tuesday")
        'VL.Add("4", "Wednesday")
        'VL.Add("5", "Thursday")
        'VL.Add("6", "Friday")
        'VL.Add("7", "Saturday")

        With grdGLTPARM3X.DisplayLayout.Bands("GLTPARM3X")
            .Columns("YEAR").Header.Fixed = True
        End With

        cbePERIOD_END_DAY.DataSource = dst.Tables("ASTPCTLP") ' TACMAIN1.CodeValueList("PERIOD_END_DAY")
        cbeWEEK_END_DAY.DataSource = dst.Tables("ASTPCTLW")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Done", "Cancel"
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            If tf Then
                If EntryMode = "E" Then
                    .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Update").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
                End If
                If EntryMode = "L" Then
                    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
                End If
            Else
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End If

            .Groups("Re-Generate Calendar").Visible = tf And (EntryMode = "E")
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
            If EntryMode = "E" Then
                grdGLTPARM2X.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.True
                grdGLTPARM3X.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.True
            Else
                grdGLTPARM2X.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False
                grdGLTPARM3X.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.False
            End If

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.Tables("GLTPARM2X").Rows.Clear()
        dst.Tables("GLTPARM3X").Rows.Clear()
        Fill_Records("ASTPCTL1", "", True, "Select * from ASTPCTL1")
        Set_Read_Only(grpASTPCTL1, True)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Calendar")

        Call Save_Header_Fields(UltraGroupBox1)

        tkbP.Value = Absx1.numFor("P01_CAL_OFFSET").Value
        Setup_P01_CAL_OFFSET()

        tkbW.Value = Val(Absx1.numFor("P01_CAL_OFFSET_YW").Value & "")
        Setup_P01_CAL_OFFSET_YW()

        Fill_Records("GLTPARM2", "", True, "Select * from GLTPARM2")
        Fill_Records("GLTPARM2X")
        Sort_grdColumns(grdGLTPARM2X, "YEAR", True)
        grdGLTPARM2X.DisplayLayout.Bands(0).CardView = True

        Fill_Records("GLTPARM3", "", True, "Select * from GLTPARM3")
        Fill_Records("GLTPARM3X")
        Sort_grdColumns(grdGLTPARM3X, "YEAR", True)
        grdGLTPARM3X.DisplayLayout.Bands(0).CardView = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()

        ASCDATA1.ExecuteSQL("Delete from GLTPARM2")
        ASCDATA1.ExecuteSQL("Delete from GLTPARM3")

        dst.Tables("GLTPARM2").Rows.Clear()
        For Each rowGLTPARM2X As DataRow In dst.Tables("GLTPARM2X").Rows
            For P As Integer = 1 To 12
                Dim rowGLTPARM2 As DataRow = dst.Tables("GLTPARM2").NewRow
                Dim YP As String = rowGLTPARM2X.Item("YEAR") & Format(P, "00")
                Dim PRD_END_DATE As Date = rowGLTPARM2X.Item("END_" & Format(P, "00"))
                rowGLTPARM2.Item("OPS_YYYYPP") = YP
                rowGLTPARM2.Item("PRD_END_DATE") = PRD_END_DATE
                rowGLTPARM2.Item("LEGEND") = Mid(YP, 1, 4) & "-" & Mid(YP, 5, 2) & " (" & Format(PRD_END_DATE, "MMM") & "'" & Format(PRD_END_DATE, "yy") & ")"
                dst.Tables("GLTPARM2").Rows.Add(rowGLTPARM2)
            Next
        Next

        dst.Tables("GLTPARM3").Rows.Clear()

        Dim WEEK_END_DAY As Integer = Absx1.cbeFor("WEEK_END_DAY").Value
        Dim P01_CAL_OFFSET As Integer = tkbP.Value
        Dim P01_CAL_OFFSET_YW As Integer = tkbW.Value
        Dim WEEK_SCHEME As Integer = Absx1.optFor("WEEK_SCHEME").Value & ""

        For Each rowGLTPARM3X As DataRow In dst.Tables("GLTPARM3X").Rows
            Dim W As Integer = 0
            For M As Integer = 1 To 12
                Dim MAX_WEEK As Integer = Val(Mid(WEEK_SCHEME, ((M - 1) Mod 3) + 1, 1))
                If M = 12 And rowGLTPARM3X.Item("END_53") & "" <> "" Then
                    MAX_WEEK = MAX_WEEK + 1
                End If
                For REL_WEEK As Integer = 1 To MAX_WEEK
                    W += 1
                    If Not rowGLTPARM3X.Item("END_" & Format(W, "00")).Equals(DBNull.Value) Then
                        Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").NewRow
                        Dim YW As String = rowGLTPARM3X.Item("YEAR") & Format(W, "00")
                        Dim YY As String = rowGLTPARM3X.Item("YEAR")
                        Dim MM As Integer = M + P01_CAL_OFFSET_YW
                        If MM <= 0 Then MM = MM + 12
                        If MM > 12 Then
                            MM = MM - 12
                            YY = YY + 1
                        End If
                        Dim YM As String = YY & Format(MM, "00")
                        Dim PP As Integer = MM - P01_CAL_OFFSET
                        If PP <= 0 Then PP = PP + 12
                        If PP > 12 Then
                            PP = PP - 12
                            YY = YY + 1
                        End If
                        Dim YP As String = YY & Format(PP, "00")
                        Dim WEEK_END_DATE As Date = rowGLTPARM3X.Item("END_" & Format(W, "00"))
                        rowGLTPARM3.Item("YYYYWW") = YW
                        rowGLTPARM3.Item("WEEK_END_DATE") = WEEK_END_DATE
                        rowGLTPARM3.Item("LEGEND") = Mid(YW, 1, 4) & "-" & Mid(YW, 5, 2) & " (" & Format(CDate(Mid(YM, 5, 2) & "/01/" & Mid(YM, 1, 4)), "MMM") & ":" & Format(REL_WEEK, "0") & "/" & Format(MAX_WEEK, "0") & ")"

                        rowGLTPARM3.Item("YYYYPP") = YP
                        rowGLTPARM3.Item("YYYYMM") = YM
                        rowGLTPARM3.Item("REL_WEEK") = REL_WEEK
                        rowGLTPARM3.Item("MAX_WEEK") = MAX_WEEK

                        dst.Tables("GLTPARM3").Rows.Add(rowGLTPARM3)
                    End If

                Next
            Next
        Next

        Update_Record_TDA("ASTPCTL1")
        Update_Record_TDA("GLTPARM2")
        Update_Record_TDA("GLTPARM3")

        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "YYYY"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Show Calendar", e)
            '    End If
        End Select
    End Sub
#End Region

    Sub Print_Report()
        'Call Print_Report_Begin()

        'Dim YYYY As Integer = Val(HFs("YYYY"))
        'CR_params.Add("Y1", Format(YYYY + 0, "0000"))
        'CR_params.Add("Y2", Format(YYYY + 1, "0000"))
        'CR_params.Add("Y3", Format(YYYY + 2, "0000"))
        'CR_params.Add("Y4", Format(YYYY + 3, "0000"))
        'CR_params.Add("Y5", Format(YYYY + 4, "0000"))
        'Generate_Report("ASRWCAL1", "Weekly Retail Calendar")

        'Call Print_Report_End()
    End Sub

    'Private Sub grdASTWCAL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
    '    If e.Row.Cells("WEEK_NO").Text = "" Then
    '        e.Row.Appearance.BackColor = Drawing.Color.AliceBlue
    '        e.Row.Height = 10
    '    End If
    'End Sub

    Private Sub tkbP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tkbP.ValueChanged
        Setup_P01_CAL_OFFSET()
    End Sub

    Sub Setup_P01_CAL_OFFSET()
        Dim P01_CAL_OFFSET As Integer = tkbP.Value
        Absx1.numFor("P01_CAL_OFFSET").Value = tkbP.Value
        Dim YP01 As String = Format(Now, "yyyy") & "01"
        Dim YM01 As String = ASCMAIN1.Period_Calc(YP01, P01_CAL_OFFSET)
        Dim YM01d As Date = CDate(Mid(YM01, 5, 2) & "/01/" & Mid(YM01, 1, 4))
        Dim LEGEND_01 As String = Format(YM01d, "MMM") & "'" & Format(YM01d, "yy")

        Dim YP12 As String = Format(Now, "yyyy") & "12"
        Dim YM12 As String = ASCMAIN1.Period_Calc(YP12, P01_CAL_OFFSET)
        Dim YM12d As Date = CDate(Mid(YM12, 5, 2) & "/01/" & Mid(YM12, 1, 4))
        Dim LEGEND_12 As String = Format(YM12d, "MMM") & "'" & Format(YM12d, "yy")

        lblP.Text = "Fiscal " & Mid(YP01, 1, 4) & " is from " & LEGEND_01 & " thru " & LEGEND_12

        grdGLTPARM2X.DisplayLayout.Bands(0).Columns("YEAR").Width = 600
        For i As Integer = 1 To 12
            Dim M As Integer = i + P01_CAL_OFFSET
            If M > 12 Then M = M - 12
            If M <= 0 Then M = M + 12
            Dim MMM As String = Format(CDate(Format(M, "00") & "/01/" & Format(Now, "yyyy")), "MMM")
            grdGLTPARM2X.DisplayLayout.Bands(0).Columns("END_" & Format(i, "00")).Header.Caption = MMM
            grdGLTPARM2X.DisplayLayout.Bands(0).Columns("END_" & Format(i, "00")).Header.Appearance.TextHAlign = HAlign.Left
            grdGLTPARM2X.DisplayLayout.Bands(0).Columns("END_" & Format(i, "00")).Width = 600
        Next
    End Sub

    Sub Setup_P01_CAL_OFFSET_YW()
        Dim P01_CAL_OFFSET_YW As Integer = tkbW.Value
        Absx1.numFor("P01_CAL_OFFSET_YW").Value = tkbW.Value
        Dim YP01 As String = Format(Now, "yyyy") & "01"
        Dim YM01 As String = ASCMAIN1.Period_Calc(YP01, P01_CAL_OFFSET_YW)
        Dim YM01d As Date = CDate(Mid(YM01, 5, 2) & "/01/" & Mid(YM01, 1, 4))
        Dim LEGEND_01 As String = Format(YM01d, "MMM") & "'" & Format(YM01d, "yy")

        Dim YP12 As String = Format(Now, "yyyy") & "12"
        Dim YM12 As String = ASCMAIN1.Period_Calc(YP12, P01_CAL_OFFSET_YW)
        Dim YM12d As Date = CDate(Mid(YM12, 5, 2) & "/01/" & Mid(YM12, 1, 4))
        Dim LEGEND_12 As String = Format(YM12d, "MMM") & "'" & Format(YM12d, "yy")

        lblW.Text = "Weekly Calendar " & Mid(YP01, 1, 4) & " is from " & LEGEND_01 & " thru " & LEGEND_12


        Dim WEEK_SCHEME As Integer = Val(Absx1.optFor("WEEK_SCHEME").Value & "")

        grdGLTPARM3X.DisplayLayout.Bands(0).Columns("YEAR").Width = 600
        Dim W As Integer = 0
        For M As Integer = 1 To 12
            Dim MAX_WEEK As Integer = Val(Mid(WEEK_SCHEME, ((M - 1) Mod 3) + 1, 1))
            For REL_WEEK As Integer = 1 To MAX_WEEK
                W += 1
                grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Header.Caption = Format(W, "00") & " (" & Format(REL_WEEK, "0") & "/" & Format(MAX_WEEK, "0") & ")"
                grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Header.Appearance.TextHAlign = HAlign.Left
                grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Width = 600

                If M Mod 2 = 1 Then
                    grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).CellAppearance.BackColor = Drawing.Color.LightBlue
                End If
            Next
        Next
        W += 1
        grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Header.Caption = Format(W, "00")
        grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Header.Appearance.TextHAlign = HAlign.Left
        grdGLTPARM3X.DisplayLayout.Bands(0).Columns("END_" & Format(W, "00")).Width = 600

    End Sub

    Sub Load_GLTPARM2X_from_GLTPARM2()

    End Sub

    Sub ReGenerate_Period_End_Dates()
        ' THIS MAY NOT WORK FOR OFFSETS > 0

        Dim PERIOD_END_DAY As Integer = Absx1.cbeFor("PERIOD_END_DAY").Value
        Dim P01_CAL_OFFSET As Integer = tkbP.Value

        dst.Tables("GLTPARM2X").Rows.Clear()
        Dim Y As Integer = Val(Format(Now, "yyyy"))
        Dim YEARS_BACK As Integer = Val(Absx1.numFor("YEARS_BACK").Value & "")
        Dim YEARS_FORWARD As Integer = Val(Absx1.numFor("YEARS_FORWARD").Value & "")
        For YEAR As Integer = Y - YEARS_BACK To Y + YEARS_FORWARD
            Dim rowGLTPARM2X As DataRow = dst.Tables("GLTPARM2X").NewRow
            rowGLTPARM2X.Item("YEAR") = Format(YEAR, "0000")
            For M As Integer = 1 To 12
                Dim PRD_END_DATE As Date = CDate(Format(M, "00") & "/01/" & Format(YEAR, "0000")).AddMonths(P01_CAL_OFFSET).AddMonths(1)
                If PERIOD_END_DAY = 0 Then
                    PRD_END_DATE = PRD_END_DATE.AddDays(-1)
                Else
                    For I As Integer = 1 To 7
                        PRD_END_DATE = PRD_END_DATE.AddDays(-1)
                        If Val(PRD_END_DATE.DayOfWeek) = PERIOD_END_DAY - 1 Then
                            Exit For
                        End If
                    Next
                End If
                rowGLTPARM2X.Item("END_" & Format(M, "00")) = PRD_END_DATE
            Next
            dst.Tables("GLTPARM2X").Rows.Add(rowGLTPARM2X)
        Next

    End Sub

    Sub ReGenerate_Week_End_Dates()

        Dim WEEK_END_DAY As Integer = Absx1.cbeFor("WEEK_END_DAY").Value    ' 7 = Sat
        Dim P01_CAL_OFFSET_YW As Integer = tkbW.Value   ' 0 = Jan-Dec, 1 = Feb-Jan, -1 Dec-Nov
        Dim WEEK_SCHEME As Integer = Absx1.optFor("WEEK_SCHEME").Value & "" ' 454

        dst.Tables("GLTPARM3X").Rows.Clear()
        Dim Y As Integer = Val(Format(Now, "yyyy"))
        Dim YEARS_BACK As Integer = Val(Absx1.numFor("YEARS_BACK").Value & "")
        Dim YEARS_FORWARD As Integer = Val(Absx1.numFor("YEARS_FORWARD").Value & "")

        Dim first_date As Boolean = True
        Dim WEEK_END_DATE As Date
        For YEAR As Integer = Y - YEARS_BACK To Y + YEARS_FORWARD
            If first_date Then
                Dim M As Integer = P01_CAL_OFFSET_YW + 1
                Dim YY As Integer = YEAR
                If M <= 0 Then
                    M = M + 12
                    YY = YY - 1
                End If
                If M >= 12 Then
                    M = M - 12
                    YY = YY + 1
                End If
                WEEK_END_DATE = CDate(Format(M, "00") & "/01/" & Format(YY, "0000"))
                Do While Val(WEEK_END_DATE.DayOfWeek) <> WEEK_END_DAY - 1
                    WEEK_END_DATE = WEEK_END_DATE.AddDays(1)
                Loop
                If WEEK_END_DATE.Day < 4 Then
                    WEEK_END_DATE = WEEK_END_DATE.AddDays(7)
                End If
                WEEK_END_DATE = WEEK_END_DATE.AddDays(-7)
                first_date = False
            End If

            Dim rowGLTPARM3X As DataRow = dst.Tables("GLTPARM3X").NewRow
            rowGLTPARM3X.Item("YEAR") = Format(YEAR, "0000")
            Dim W As Integer = 0
            For M As Integer = 1 To 12
                Dim MAX_WEEK As Integer = Val(Mid(WEEK_SCHEME, ((M - 1) Mod 3) + 1, 1))
                'If M = 12 And Format(WEEK_END_DATE.AddDays(7), "MM") = Format(WEEK_END_DATE.AddDays(7 + 7 * MAX_WEEK), "MM") Then
                If M = 12 And WEEK_END_DATE.AddDays(7 + 7 * MAX_WEEK).Day < 4 Then
                    MAX_WEEK = MAX_WEEK + 1
                End If
                For REL_WEEK As Integer = 1 To MAX_WEEK
                    WEEK_END_DATE = WEEK_END_DATE.AddDays(7)
                    W += 1
                    rowGLTPARM3X.Item("END_" & Format(W, "00")) = WEEK_END_DATE
                Next
            Next
            dst.Tables("GLTPARM3X").Rows.Add(rowGLTPARM3X)
        Next

    End Sub

    Private Sub cmdGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerate.Click

        ReGenerate_Period_End_Dates()
        ReGenerate_Week_End_Dates()

    End Sub

    Private Sub grdGLTPARM3X_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTPARM3X.InitializeLayout

    End Sub

    Private Sub grdGLTPARM3X_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTPARM3X.InitializeRow

    End Sub

    Private Sub tkbW_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tkbW.ValueChanged
        Setup_P01_CAL_OFFSET_YW()
    End Sub
End Class