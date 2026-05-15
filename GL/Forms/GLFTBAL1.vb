Public Class GLFTBAL1

    Dim ACCT_END_BAL_TOTAL As Decimal = 0
    Dim GLTTBAL1 As String = ""
    Dim sqlGLTSUMJ2 As String = ""
    Dim GLTACCT3 As String
    Dim GLTACCT2 As String
    Dim GLTACCT4 As String

    ' SET UP AHA PROJECT BY DOING A CLONE FROM ABSCODERS

    ' HIDE ACCT_YEAR, ACCT_BEG_BAL
    ' HEADINGS FOR SEG2/3/4
    ' PRETTY HEADINGS For ACCT_CODE, ACCT_DESC
    ' PRETTY HEADINGS For Each Of THE PERIODS (Jan'20 instead of ACCT_ACT_PO1)
    ' PIN THE COLS BEFORE THE 1ST MON

    ' Do THE BUDGETS TAB - REUSING AS MUCH CODE AS POSSIBLE


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Create_Work_Table("000000")

        With dst
            With .Tables.Add("GLTTBAL0")
                With .Columns
                    .Add("ACCT_TYPE", GetType(System.String))
                    .Add("ACCT_TYPE_DESC", GetType(System.String))
                    .Add("ACCT_BEG_BAL", GetType(System.Decimal))
                    .Add("ACCT_DR", GetType(System.Decimal))
                    .Add("ACCT_CR", GetType(System.Decimal))
                    .Add("ACCT_END_BAL", GetType(System.Decimal))
                    .Add("ACCT_TRANS", GetType(System.Int32))
                    .Add("ACCT_TYPE_SEQ", GetType(System.Int32))
                    For P As Integer = 1 To 12
                        .Add("P" & Format(P, "00"), GetType(System.Decimal))
                    Next
                End With
                .PrimaryKey = New DataColumn() {.Columns("ACCT_TYPE")}
            End With

            Dim sqlP12 As String = ""
            For P As Integer = 1 To 12
                sqlP12 &= ", GLTACCT3.ACCT_ACT_P" & Format(P, "00") & " " & "P" & Format(P, "00") & vbCrLf
            Next


            ASCMAIN1.sql = "Select GLTACCT3.ACCT_CODE" & vbCrLf _
                & ", GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" & vbCrLf _
                & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" & vbCrLf _
                & ", GLTACCT3.ACCT_BEG_BAL" & vbCrLf _
                & ", ACCT_BEG_BAL ACCT_END_BAL " & vbCrLf _
                & sqlP12 _
                & " from GLTACCT3,GLTACCT1 " & vbCrLf _
                & " where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE "
            Create_TDA(.Tables.Add, "GLTTBAL1", "**", 0, False, "", 4)
            With .Tables("GLTTBAL1").Columns
                .Add("ACCT_DR", GetType(System.Decimal))
                .Add("ACCT_CR", GetType(System.Decimal))
                .Add("ACCT_TRANS", GetType(System.Int32))
                'For P As Integer = 1 To 12
                '    .Add("P" & Format(P, "00"), GetType(System.Decimal))
                'Next
            End With

            Create_Relation("GLTTBAL0", "GLTTBAL1", "ACCT_TYPE")

            With .Tables("GLTTBAL0")
                .Columns("ACCT_BEG_BAL").Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).ACCT_BEG_BAL)"
                .Columns("ACCT_DR").Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).ACCT_DR)"
                .Columns("ACCT_CR").Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).ACCT_CR)"
                .Columns("ACCT_END_BAL").Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).ACCT_END_BAL)"
                .Columns("ACCT_TRANS").Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).ACCT_TRANS)"
                For P As Integer = 1 To 12
                    .Columns("P" & Format(P, "00")).Expression = "SUM(Child(GLTTBAL0_GLTTBAL1).P" & Format(P, "00") & ")"
                Next
            End With


            ASCMAIN1.sql = "Select Distinct GLTJRNL1.JOURNAL_TYPE, GLTTYPE1.JOURNAL_TYPE_DESC" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1,GLTTYPE1" & vbCrLf _
                & " where GLTTYPE1.JOURNAL_TYPE = GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP >= :PARM1 and GLTDETL1.OPS_YYYYPP <= :PARM2" & vbCrLf _
                & " group by GLTJRNL1.JOURNAL_TYPE, GLTTYPE1.JOURNAL_TYPE_DESC"
            Create_TDA(.Tables.Add, "GLTSUMJ1", "**", 0, False, "VV", 1)
            With .Tables("GLTSUMJ1")
                With .Columns
                    .Add("ACCT_DR", GetType(System.Decimal))
                    .Add("ACCT_CR", GetType(System.Decimal))
                    .Add("ACCT_TRANS", GetType(System.Int32))
                    For P As Integer = 1 To 12
                        .Add("P" & Format(P, "00"), GetType(System.Decimal))
                    Next
                End With
                .PrimaryKey = New DataColumn() {.Columns("JOURNAL_TYPE")}
            End With


            Dim sqlP12j As String = ""
            Dim sqlP12j0 As String = ", Sum(DECODE(SUBSTR(GLTDETL1.OPS_YYYYPP,5,2),'??',GLTDETL1.DETL_POSTING_AMT,0)) P??"
            For P As Integer = 1 To 12
                 sqlP12j &= Replace(sqlP12j0, "??", Format(P, "00")) & vbCrLf
            Next

            ASCMAIN1.sql = "Select GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE" & vbCrLf _
                & ", GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" & vbCrLf _
                & ", GLTACCT1.ACCT_DESC" & vbCrLf _
                & ", Sum (DECODE(GLTDETL1.OPS_YYYYPP,'000000',CASE WHEN DETL_POSTING_AMT > 0 THEN DETL_POSTING_AMT ELSE 0 END,0)) ACCT_DR" & vbCrLf _
                & ", Sum (DECODE(GLTDETL1.OPS_YYYYPP,'000000',CASE WHEN DETL_POSTING_AMT < 0 THEN -1 * DETL_POSTING_AMT ELSE 0 END,0)) ACCT_CR" & vbCrLf _
                & ", Sum (DECODE(GLTDETL1.OPS_YYYYPP,'000000',1,0)) ACCT_TRANS" & vbCrLf _
                & sqlP12j _
                & " from GLTDETL1,GLTJRNL1,GLTACCT1 " & vbCrLf _
                & " where GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE " & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP >= '000001' and GLTDETL1.OPS_YYYYPP <= '000000'" & vbCrLf _
                & " group by GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE" & vbCrLf _
                & ", GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" & vbCrLf _
                & ", GLTACCT1.ACCT_DESC"
            sqlGLTSUMJ2 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "GLTSUMJ2", "**", 0, False, "VVVVV", 5)

            With .Tables("GLTSUMJ2")
                .Columns("ACCT_TRANS").DataType = GetType(System.Int32)
            End With

            Create_Relation("GLTSUMJ1", "GLTSUMJ2", "JOURNAL_TYPE")

            With .Tables("GLTSUMJ1")
                .Columns("ACCT_DR").Expression = "SUM(Child(GLTSUMJ1_GLTSUMJ2).ACCT_DR)"
                .Columns("ACCT_CR").Expression = "SUM(Child(GLTSUMJ1_GLTSUMJ2).ACCT_CR)"
                .Columns("ACCT_TRANS").Expression = "SUM(Child(GLTSUMJ1_GLTSUMJ2).ACCT_TRANS)"
                For P As Integer = 1 To 12
                    .Columns("P" & Format(P, "00")).Expression = "SUM(Child(GLTSUMJ1_GLTSUMJ2).P" & Format(P, "00") & ")"
                Next
            End With

            ASCMAIN1.sql = "Select * from " & GLTACCT3
            Create_TDA(.Tables.Add, "GLTACCT3", "**", 0, False, "", 0)
            With .Tables("GLTACCT3")
                Dim TOTAL As String = ""
                For P As Integer = 1 To 12
                    TOTAL &= "+ISNULL(ACCT_ACT_P" & Format(P, "00") & ", 0)"
                Next P
                .Columns.Add("TOTAL", GetType(System.Decimal), Mid(TOTAL, 2))
            End With


            ASCMAIN1.sql = "Select * from " & GLTACCT2
            Create_TDA(.Tables.Add, "GLTACCT2", "**", 0, False, "", 0)
            With .Tables("GLTACCT2")
                Dim TOTAL As String = ""
                For P As Integer = 1 To 12
                    TOTAL &= "+ISNULL(ACCT_BUD_P" & Format(P, "00") & ", 0)"
                Next P
                .Columns.Add("TOTAL", GetType(System.Decimal), Mid(TOTAL, 2))
            End With

            ASCMAIN1.sql = "Select * from " & GLTACCT4
            Create_TDA(.Tables.Add, "GLTACCT4", "**", 0, False, "", 0)
            With .Tables("GLTACCT4")
                Dim TOTAL As String = ""
                For P As Integer = 1 To 12
                    TOTAL &= "+ISNULL(ACCT_BUD_P" & Format(P, "00") & ", 0)"
                Next P
                .Columns.Add("TOTAL", GetType(System.Decimal), Mid(TOTAL, 2))
            End With



        End With

        grdGLTTBAL1.DataSource = dst.Tables("GLTTBAL0")
        grdGLTSUMJ1.DataSource = dst.Tables("GLTSUMJ1")
        grdGLTACCT3.DataSource = dst.Tables("GLTACCT3")
        grdGLTACCT2.DataSource = dst.Tables("GLTACCT2")
        grdGLTACCT4.DataSource = dst.Tables("GLTACCT4")

        Add_ACCT_TYPEs("GLTTBAL0", True)

        Set_SEGS(grdGLTTBAL1, "GLTTBAL0_GLTTBAL1")
        Set_SEGS(grdGLTSUMJ1, "GLTSUMJ1_GLTSUMJ2")

        grdGLTSUMJ1.DisplayLayout.Bands("GLTSUMJ1_GLTSUMJ2").SummaryFooterCaption = "Totals For Journal Type: [JOURNAL_TYPE] [JOURNAL_TYPE_DESC]"

        Create_Summary(grdGLTTBAL1, New String() {"ACCT_BEG_BAL", "ACCT_DR", "ACCT_CR", "ACCT_END_BAL", "ACCT_TRANS", "P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "P10", "P11", "P12"})
        Create_Summary(grdGLTTBAL1, New String() {"ACCT_BEG_BAL", "ACCT_DR", "ACCT_CR", "ACCT_END_BAL", "ACCT_TRANS", "P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "P10", "P11", "P12"}, , "GLTTBAL0_GLTTBAL1")
        Create_Summary(grdGLTTBAL1, "ACCT_CODE", "Count", "GLTTBAL0_GLTTBAL1")

        Create_Summary(grdGLTSUMJ1, New String() {"ACCT_DR", "ACCT_CR", "ACCT_TRANS", "P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "P10", "P11", "P12"})
        Create_Summary(grdGLTSUMJ1, New String() {"ACCT_DR", "ACCT_CR", "ACCT_TRANS", "P01", "P02", "P03", "P04", "P05", "P06", "P07", "P08", "P09", "P10", "P11", "P12"}, , "GLTSUMJ1_GLTSUMJ2")
        Create_Summary(grdGLTSUMJ1, "ACCT_CODE", "Count", "GLTSUMJ1_GLTSUMJ2")

        With grdGLTTBAL1.DisplayLayout
            For B As Integer = 0 To 1
                With .Bands(B)
                    If B = 0 Then
                        For Each COLUMN_NAME As String In New String() _
                            {"ACCT_TYPE", "ACCT_TYPE_DESC"}
                            With .Columns(COLUMN_NAME)
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        Next
                    Else
                        For Each COLUMN_NAME As String In New String() _
                            {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}
                            With .Columns(COLUMN_NAME)
                                .Width = IIf(COLUMN_NAME = "ACCT_DESC", 250, 60)
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        Next
                    End If
                    For Each COLUMN_NAME As String In New String() _
                        {"ACCT_BEG_BAL", "ACCT_DR", "ACCT_CR", "ACCT_END_BAL", "ACCT_TRANS"}
                        With .Columns(COLUMN_NAME)
                            .Width = IIf(COLUMN_NAME = "ACCT_TRANS", 60, 110)
                            .Header.Appearance.BackColor = Drawing.Color.White
                            .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        End With
                    Next
                    For P As Integer = 1 To 12
                        With .Columns("P" & Format(P, "00"))
                            .Width = 110
                            .Header.Appearance.BackColor = Drawing.Color.White
                            .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        End With
                    Next
                End With
            Next
        End With

        With grdGLTSUMJ1.DisplayLayout
            For B As Integer = 0 To 1
                With .Bands(B)
                    If B = 0 Then
                        For Each COLUMN_NAME As String In New String() _
                            {"JOURNAL_TYPE", "JOURNAL_TYPE_DESC"}
                            With .Columns(COLUMN_NAME)
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        Next
                    Else
                        For Each COLUMN_NAME As String In New String() _
                            {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}
                            With .Columns(COLUMN_NAME)
                                .Width = IIf(COLUMN_NAME = "ACCT_DESC", 250, 60)
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        Next
                    End If
                    For Each COLUMN_NAME As String In New String() _
                        {"ACCT_DR", "ACCT_CR", "ACCT_TRANS"}
                        With .Columns(COLUMN_NAME)
                            .Width = IIf(COLUMN_NAME = "ACCT_TRANS", 60, 110)
                            .Header.Appearance.BackColor = Drawing.Color.White
                            .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        End With
                    Next
                    For P As Integer = 1 To 12
                        With .Columns("P" & Format(P, "00"))
                            .Width = 110
                            .Header.Appearance.BackColor = Drawing.Color.White
                            .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        End With
                    Next
                End With
            Next
        End With

        Breakout_By()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Report"
                Call Print_Report_Begin()
                CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
                CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
                CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
                CR_params.Add("CHKSEG2", "1")
                CR_params.Add("CHKSEG3", "1")
                CR_params.Add("CHKSEG4", "1")
                CR_params.Add("GYPLEGEND", "???")
                Generate_Report("GLRTBAL1")
                Call Print_Report_End()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If



    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTTBAL1", "GLTSUMJ1", "GLTSUMJ2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("OPS_YYYYPP").Text = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Account Summary Data")

        grdGLTTBAL1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim ACCT_YEAR As String = Mid$(HFs("OPS_YYYYPP"), 1, 4)
        Dim P As Integer = Val(Mid$(HFs("OPS_YYYYPP"), 5, 2))

        Dim sql_BEG_BAL As String = ""
        Dim sql_END_BAL As String = ""
        For i As Integer = 1 To P
            Dim z As String = " + NVL(GLTACCT3.ACCT_ACT_P" & Format(i, "00") & ",0)"
            If i < P Then
                sql_BEG_BAL = sql_BEG_BAL & z
            End If
            sql_END_BAL = sql_END_BAL & z
        Next

        Dim GLTACCT3 As String = GL_Prep(ACCT_YEAR, ACCT_YEAR)

        Dim sql As String = ""
        sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE " _
            & ", Sum (CASE WHEN DETL_POSTING_AMT > 0 THEN DETL_POSTING_AMT ELSE 0 END) ACCT_DR" _
            & ", Sum (CASE WHEN DETL_POSTING_AMT < 0 THEN -1 * DETL_POSTING_AMT ELSE 0 END) ACCT_CR" _
            & ", Count (*) ACCT_TRANS" _
            & " from GLTDETL1 where OPS_YYYYPP = '" & HFs("OPS_YYYYPP") & "'" _
            & " group by ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"
        Dim GLTACCT3x As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTACCT3x & " Add Primary Key (ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE)")
        ASCMAIN1.AnalyzeTable(GLTACCT3x)

        Dim sqlP12 As String = ""
        For PX As Integer = 1 To 12
            sqlP12 &= ", sum (GLTACCT3.ACCT_ACT_P" & Format(PX, "00") & ") " & "P" & Format(PX, "00") & vbCrLf
        Next

        sql = "Select GLTACCT3.ACCT_CODE" & vbCrLf _
            & ", " & IIf(chkSEG2_CODE.Checked, "GLTACCT3.SEG2_CODE", "'000'") & " SEG2_CODE" & vbCrLf _
            & ", " & IIf(chkSEG3_CODE.Checked, "GLTACCT3.SEG3_CODE", "'000'") & " SEG3_CODE" & vbCrLf _
            & ", " & IIf(chkSEG4_CODE.Checked, "GLTACCT3.SEG4_CODE", "'000'") & " SEG4_CODE" & vbCrLf _
            & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" _
            & ", Sum (GLTACCT3.ACCT_BEG_BAL" & sql_BEG_BAL & ") ACCT_BEG_BAL" & vbCrLf _
            & ", Sum (GLTACCT3.ACCT_BEG_BAL" & sql_END_BAL & ") ACCT_END_BAL" & vbCrLf _
            & sqlP12 _
            & ", Sum (GLTACCT3X.ACCT_DR) ACCT_DR" & vbCrLf _
            & ", Sum (GLTACCT3X.ACCT_CR) ACCT_CR" & vbCrLf _
            & ", Sum (GLTACCT3X.ACCT_TRANS) ACCT_TRANS" & vbCrLf _
            & " from " & GLTACCT3 & " GLTACCT3,GLTACCT1," & GLTACCT3x & " GLTACCT3X" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE " & vbCrLf _
            & "   and GLTACCT3.ACCT_YEAR = '" & ACCT_YEAR & "'" & vbCrLf _
            & " and GLTACCT3.ACCT_CODE = GLTACCT3X.ACCT_CODE (+)" & vbCrLf _
            & " and GLTACCT3.SEG2_CODE = GLTACCT3X.SEG2_CODE (+)" & vbCrLf _
            & " and GLTACCT3.SEG3_CODE = GLTACCT3X.SEG3_CODE (+)" & vbCrLf _
            & " and GLTACCT3.SEG4_CODE = GLTACCT3X.SEG4_CODE (+)" & vbCrLf _
            & " group by GLTACCT3.ACCT_CODE" & vbCrLf _
            & IIf(chkSEG2_CODE.Checked, ", GLTACCT3.SEG2_CODE", "") & vbCrLf _
            & IIf(chkSEG3_CODE.Checked, ", GLTACCT3.SEG3_CODE", "") & vbCrLf _
            & IIf(chkSEG4_CODE.Checked, ", GLTACCT3.SEG4_CODE", "") & vbCrLf _
            & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC"

        GLTTBAL1 = ASCMAIN1.Temp_Table(sql)


        Set_SelectCommand("GLTTBAL1", sql)

        Fill_Records("GLTTBAL1")

        Fill_Records("GLTSUMJ1", New String() {ACCT_YEAR & "01", HFs("OPS_YYYYPP")})

        ASCMAIN1.sql = Replace(Replace(sqlGLTSUMJ2, "000000", HFs("OPS_YYYYPP")), "000001", ACCT_YEAR & "01")
        If Not chkSEG2_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG2_CODE", ", '000' SEG2_CODE", , 1)
        If Not chkSEG3_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG3_CODE", ", '000' SEG3_CODE", , 1)
        If Not chkSEG4_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG4_CODE", ", '000' SEG4_CODE", , 1)
        If Not chkSEG2_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG2_CODE", "", , 1)
        If Not chkSEG3_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG3_CODE", "", , 1)
        If Not chkSEG4_CODE.Checked Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, ", GLTDETL1.SEG4_CODE", "", , 1)

        grdGLTSUMJ1.DisplayLayout.Bands(1).Columns("SEG2_CODE").Hidden = Not chkSEG2_CODE.Checked
        grdGLTSUMJ1.DisplayLayout.Bands(1).Columns("SEG3_CODE").Hidden = Not chkSEG3_CODE.Checked
        grdGLTSUMJ1.DisplayLayout.Bands(1).Columns("SEG4_CODE").Hidden = Not chkSEG4_CODE.Checked

        Fill_Records("GLTSUMJ2", "", True, ASCMAIN1.sql)


        grdGLTTBAL1.DisplayLayout.Bands(1).Columns("SEG2_CODE").Hidden = Not chkSEG2_CODE.Checked
        grdGLTTBAL1.DisplayLayout.Bands(1).Columns("SEG3_CODE").Hidden = Not chkSEG3_CODE.Checked
        grdGLTTBAL1.DisplayLayout.Bands(1).Columns("SEG4_CODE").Hidden = Not chkSEG4_CODE.Checked


        EnforceConstraints(True)

        Sort_grdColumns(grdGLTTBAL1, "ACCT_TYPE_SEQ", False, 0)
        Sort_grdColumns(grdGLTTBAL1, "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE", False, 1)
        Sort_grdColumns(grdGLTSUMJ1, "JOURNAL_TYPE", False, 0)
        Sort_grdColumns(grdGLTSUMJ1, "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE", False, 1)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGLTTBAL1, grdGLTSUMJ1}
            With grd.DisplayLayout
                For B As Integer = 0 To 1
                    With .Bands(B)
                        For PX As Integer = 1 To 12
                            With .Columns("P" & Format(PX, "00"))
                                Dim LEGEND As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(PX, "00"))
                                .Header.Caption = Mid(LEGEND, 10, 6)
                                .Hidden = (PX > P)
                            End With
                        Next
                    End With
                Next

                Dim X As Integer = 1
                If chkSEG2_CODE.Checked Then X += 1
                If chkSEG3_CODE.Checked Then X += 1
                If chkSEG4_CODE.Checked Then X += 1

                Dim COLUMN_NAME_cs As String = "ACCT_TYPE"
                If grd.Name = "grdGLTSUMJ1" Then
                    COLUMN_NAME_cs = "JOURNAL_TYPE"
                End If
                .Bands(0).Columns(COLUMN_NAME_cs).ColSpan = X
                .Bands(0).Columns(COLUMN_NAME_cs).Width = X * 60 + 10

                With .Bands(1)
                    For Each COLUMN_NAME As String In New String() _
                                {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}
                        With .Columns(COLUMN_NAME)
                            .Width = 60
                        End With
                    Next
                End With
            End With
        Next


        '    If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
        UltraTabControl1.Tabs("Export").Visible = True
        ASCMAIN1.sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE" & vbCrLf _
& ", ACCT_CODE || '-' || SEG2_CODE || '-' || SEG3_CODE || '-' || SEG4_CODE ACCT" & vbCrLf _
& ", ACCT_DESC, ACCT_BEG_BAL, ACCT_DR, ACCT_CR, NVL(ACCT_DR,0) + NVL(ACCT_DR,0) ACCT_NET, ACCT_END_BAL" & vbCrLf _
& " from " & GLTTBAL1
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        grdTBExport.DataSource = tbl
        Sort_grdColumns(grdTBExport, "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")
        With grdTBExport.DisplayLayout.Bands(0)
            .Columns("ACCT_CODE").Header.Caption = "Acct"
            .Columns("SEG2_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
            .Columns("SEG3_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
            .Columns("SEG4_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
            .Columns("ACCT_DESC").Header.Caption = "Description"
            .Columns("ACCT_BEG_BAL").Header.Caption = "Beg Bal"
            .Columns("ACCT_DR").Header.Caption = "DR"
            .Columns("ACCT_CR").Header.Caption = "CR"
            .Columns("ACCT_NET").Header.Caption = "Net"
            .Columns("ACCT_END_BAL").Header.Caption = "End Bal"
        End With


        '    Else
        '        UltraTabControl1.Tabs("Export").Visible = False

        '    End If


        Create_Work_Table(HFs("OPS_YYYYPP"))

        Fill_Records("GLTACCT3")
        Sort_grdColumns(grdGLTACCT3, "ACCT_CODE")
        grdGLTACCT3.DisplayLayout.Bands(0).Columns("ACCT_YEAR").Hidden = True
        grdGLTACCT3.DisplayLayout.Bands(0).Columns("ACCT_BEG_BAL").Hidden = True
        grdGLTACCT3.DisplayLayout.Bands(0).Columns("ACCT_ACT_P13").Hidden = True

        With grdGLTACCT3.DisplayLayout.Bands(0)
            .Columns("ACCT_CODE").Header.Fixed = True
            .Columns("SEG2_CODE").Header.Fixed = True
            .Columns("SEG3_CODE").Header.Fixed = True
            .Columns("SEG4_CODE").Header.Fixed = True
            .Columns("ACCT_DESC").Header.Fixed = True


            .Columns("ACCT_CODE").Header.Caption = "Acct"
            .Columns("SEG2_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
            .Columns("SEG3_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
            .Columns("SEG4_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
            .Columns("ACCT_DESC").Header.Caption = "Description"
            .Columns("ACCT_BEG_BAL").Header.Caption = "Beg Bal"
            .Columns("TOTAL").Header.Caption = "Total"


            For PX As Integer = 1 To 12
                With .Columns("ACCT_ACT_P" & Format(PX, "00"))
                    Dim LEGEND As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(PX, "00"))
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    '   .Hidden = (PX > P)
                End With
            Next


        End With

        Fill_Records("GLTACCT2")
        Sort_grdColumns(grdGLTACCT2, "ACCT_CODE")
        grdGLTACCT2.DisplayLayout.Bands(0).Columns("ACCT_YEAR").Hidden = True
        grdGLTACCT2.DisplayLayout.Bands(0).Columns("ACCT_BEG_BAL").Hidden = True
        grdGLTACCT2.DisplayLayout.Bands(0).Columns("ACCT_BUD_P13").Hidden = True

        With grdGLTACCT2.DisplayLayout.Bands(0)
                .Columns("ACCT_CODE").Header.Fixed = True
                .Columns("SEG2_CODE").Header.Fixed = True
                .Columns("SEG3_CODE").Header.Fixed = True
                .Columns("SEG4_CODE").Header.Fixed = True
                .Columns("ACCT_DESC").Header.Fixed = True


                .Columns("ACCT_CODE").Header.Caption = "Acct"
                .Columns("SEG2_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
                .Columns("SEG3_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
                .Columns("SEG4_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
                .Columns("ACCT_DESC").Header.Caption = "Description"
                .Columns("ACCT_BEG_BAL").Header.Caption = "Beg Bal"
            .Columns("TOTAL").Header.Caption = "Total"


            For PX As Integer = 1 To 12
                With .Columns("ACCT_BUD_P" & Format(PX, "00"))
                    Dim LEGEND As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(PX, "00"))
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    '  .Hidden = (PX > P)
                End With
            Next


        End With

        Fill_Records("GLTACCT4")
        Sort_grdColumns(grdGLTACCT4, "ACCT_CODE")
        grdGLTACCT4.DisplayLayout.Bands(0).Columns("ACCT_YEAR").Hidden = True
        grdGLTACCT4.DisplayLayout.Bands(0).Columns("ACCT_BEG_BAL").Hidden = True
        grdGLTACCT4.DisplayLayout.Bands(0).Columns("ACCT_BUD_P13").Hidden = True

        With grdGLTACCT4.DisplayLayout.Bands(0)
            .Columns("ACCT_CODE").Header.Fixed = True
            .Columns("SEG2_CODE").Header.Fixed = True
            .Columns("SEG3_CODE").Header.Fixed = True
            .Columns("SEG4_CODE").Header.Fixed = True
            .Columns("ACCT_DESC").Header.Fixed = True


            .Columns("ACCT_CODE").Header.Caption = "Acct"
            .Columns("SEG2_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
            .Columns("SEG3_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
            .Columns("SEG4_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
            .Columns("ACCT_DESC").Header.Caption = "Description"
            .Columns("ACCT_BEG_BAL").Header.Caption = "Beg Bal"
            .Columns("TOTAL").Header.Caption = "Total"


            For PX As Integer = 1 To 12
                With .Columns("ACCT_BUD_P" & Format(PX, "00"))
                    Dim LEGEND As String = ASCMAIN1.Get_Legend(ACCT_YEAR & Format(PX, "00"))
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    '  .Hidden = (PX > P)
                End With
            Next



        End With




        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTTBAL1, "B", "Account Inquiry")
        Load_Popup_Menu(grdGLTSUMJ1, "B", "Account Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdX"


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case ""


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Account Inquiry"
                If grd.ActiveRow.Band.Index = 1 Then
                    Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Value
                    Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                    If rowGLTACCT1 IsNot Nothing Then
                        Context_Launch("Load", ACCT_CODE, e.Tool.Key, "GLFACTI1")
                    End If
                End If

        End Select
    End Sub

#End Region

    Private Sub chkSEG2_CODE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSEG2_CODE.CheckedChanged
        RECYCLE_VIEW()
    End Sub

    Private Sub chkSEG3_CODE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSEG3_CODE.CheckedChanged
        RECYCLE_VIEW()
    End Sub

    Private Sub chkSEG4_CODE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSEG4_CODE.CheckedChanged
        RECYCLE_VIEW()
    End Sub

    Sub RECYCLE_VIEW()
        If Me.SELECTION_NO = 0 Then Exit Sub
        If Not ScreenMode Then Exit Sub
        Click_Command("Done")
        Click_Command("View")
    End Sub

    Sub Create_Work_Table(YP As String)

        Dim initialize As Boolean = (YP = "000000")

        ASCMAIN1.sql = "Select GLTACCT3.*, GLTACCT1.ACCT_DESC" _
            & " from GLTACCT3, GLTACCT1" _
            & " where GLTACCT3.ACCT_YEAR = '" & Mid(YP, 1, 4) & "'" _
            & "   and GLTACCT1.ACCT_CODE = GLTACCT3.ACCT_CODE" _
            & "   and GLTACCT1.ACCT_TYPE In ('I','X')"

        If GLTACCT3 <> "" Then
            ASCDATA1.ExecuteSQL("Delete from " & GLTACCT3)
            ASCDATA1.ExecuteSQL("Insert into " & GLTACCT3 & " " & ASCMAIN1.sql)
        Else
            GLTACCT3 = ASCMAIN1.Temp_Table
        End If

        If Not initialize Then
            If Mid(YP, 5, 2) < "12" Then
                For P As Integer = Val(Mid(YP, 5, 2)) To 12
                    ASCDATA1.ExecuteSQL("Update " & GLTACCT3 & " set ACCT_ACT_P" & Format(P, "00") & " = 0")
                Next P
            End If
        End If


        ASCMAIN1.sql = "Select GLTACCT2.*, GLTACCT1.ACCT_DESC" _
             & " from GLTACCT2, GLTACCT1" _
             & " where GLTACCT2.ACCT_YEAR = '" & Mid(YP, 1, 4) & "'" _
             & "   and GLTACCT1.ACCT_CODE = GLTACCT2.ACCT_CODE" _
             & "   and GLTACCT1.ACCT_TYPE in ('I','X')"

        If GLTACCT2 <> "" Then
            ASCDATA1.ExecuteSQL("Delete from " & GLTACCT2)
            ASCDATA1.ExecuteSQL("Insert into " & GLTACCT2 & " " & ASCMAIN1.sql)
        Else
            GLTACCT2 = ASCMAIN1.Temp_Table
        End If

        If Not initialize Then
            If Mid(YP, 5, 2) < "12" Then
                For P As Integer = Val(Mid(YP, 5, 2)) To 12
                    ASCDATA1.ExecuteSQL("Update " & GLTACCT2 & " set ACCT_BUD_P" & Format(P, "00") & " = 0")
                Next P
            End If
        End If

        ASCMAIN1.sql = "Select GLTACCT4.*, GLTACCT1.ACCT_DESC" _
            & " from GLTACCT4, GLTACCT1" _
            & " where GLTACCT4.ACCT_YEAR = '" & Mid(YP, 1, 4) & "'" _
            & "   and GLTACCT1.ACCT_CODE = GLTACCT4.ACCT_CODE" _
            & "   and GLTACCT1.ACCT_TYPE in ('I','X')"

        If GLTACCT4 <> "" Then
            ASCDATA1.ExecuteSQL("Delete from " & GLTACCT4)
            ASCDATA1.ExecuteSQL("Insert into " & GLTACCT4 & " " & ASCMAIN1.sql)
        Else
            GLTACCT4 = ASCMAIN1.Temp_Table
        End If

        If Not initialize Then
            If Mid(YP, 5, 2) < "12" Then
                For P As Integer = Val(Mid(YP, 5, 2)) To 12
                    ASCDATA1.ExecuteSQL("Update " & GLTACCT4 & " set ACCT_BUD_P" & Format(P, "00") & " = 0")
                Next P
            End If
        End If

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged

        With UltraExplorerBar1
            If e.Tab.Key = "P&&L Actuals" Or e.Tab.Key = "P&&L Budgets-Revised" Or e.Tab.Key = "P&&L Budgets-Original" Then
                .Groups("Break-Out By").Visible = False
            Else
                .Groups("Break-Out By").Visible = True
            End If

        End With
    End Sub
End Class