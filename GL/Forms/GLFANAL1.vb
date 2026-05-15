Public Class GLFANAL1

    Dim ACCT_END_BAL_TOTAL As Double = 0

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            .Tables.Add("GLTTBAL0")
            .Tables("GLTTBAL0").Columns.Add("ACCT_TYPE", GetType(System.String))
            .Tables("GLTTBAL0").Columns.Add("ACCT_TYPE_DESC", GetType(System.String))
            .Tables("GLTTBAL0").Columns.Add("ACCT_BEG_BAL", GetType(System.Double))
            .Tables("GLTTBAL0").Columns.Add("ACCT_DR", GetType(System.Double))
            .Tables("GLTTBAL0").Columns.Add("ACCT_CR", GetType(System.Double))
            .Tables("GLTTBAL0").Columns.Add("ACCT_END_BAL", GetType(System.Double))
            .Tables("GLTTBAL0").Columns.Add("ACCT_TRANS", GetType(System.Int32))
            .Tables("GLTTBAL0").Columns.Add("ACCT_TYPE_SEQ", GetType(System.Int32))
            .Tables("GLTTBAL0").PrimaryKey = New DataColumn() {.Tables("GLTTBAL0").Columns("ACCT_TYPE")}

            ASCMAIN1.sql = "Select GLTACCT3.ACCT_CODE" _
                & ", GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
                & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" _
                & ", GLTACCT3.ACCT_BEG_BAL" _
                & ", ACCT_BEG_BAL ACCT_END_BAL " _
                & " from GLTACCT3,GLTACCT1 " _
                & " where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE "
            Create_TDA(.Tables.Add, "GLTTBAL1", "**", 0, False, "", 4)
            .Tables("GLTTBAL1").Columns.Add("ACCT_DR", GetType(System.Double))
            .Tables("GLTTBAL1").Columns.Add("ACCT_CR", GetType(System.Double))
            .Tables("GLTTBAL1").Columns.Add("ACCT_TRANS", GetType(System.Int32))

            .Relations.Add("GLTTBAL1", _
            .Tables("GLTTBAL0").Columns("ACCT_TYPE"), _
            .Tables("GLTTBAL1").Columns("ACCT_TYPE"))


            ASCMAIN1.sql = "Select Distinct GLTJRNL1.JOURNAL_TYPE, GLTTYPE1.JOURNAL_TYPE_DESC" _
            & " from GLTDETL1,GLTJRNL1,GLTTYPE1" _
            & " where GLTTYPE1.JOURNAL_TYPE = GLTJRNL1.JOURNAL_TYPE" _
            & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" _
            & "   and GLTDETL1.OPS_YYYYPP = :PARM1" _
            & " group by GLTJRNL1.JOURNAL_TYPE, GLTTYPE1.JOURNAL_TYPE_DESC"
            Create_TDA(.Tables.Add, "GLTSUMJ1", "**", 0, False, "V", 1)

            '.Tables.Add("GLTSUMJ1")
            '.Tables("GLTSUMJ1").Columns.Add("JOURNAL_TYPE", GetType(System.String))
            '.Tables("GLTSUMJ1").Columns.Add("JOURNAL_TYPE_DESC", GetType(System.String))
            .Tables("GLTSUMJ1").Columns.Add("ACCT_DR", GetType(System.Double))
            .Tables("GLTSUMJ1").Columns.Add("ACCT_CR", GetType(System.Double))
            .Tables("GLTSUMJ1").Columns.Add("ACCT_TRANS", GetType(System.Int32))
            .Tables("GLTSUMJ1").PrimaryKey = New DataColumn() {.Tables("GLTSUMJ1").Columns("JOURNAL_TYPE")}

            ASCMAIN1.sql = "Select GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE" _
            & ", GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" _
            & ", GLTACCT1.ACCT_DESC" _
            & ", Sum (CASE WHEN DETL_POSTING_AMT > 0 THEN DETL_POSTING_AMT ELSE 0 END) ACCT_DR" _
            & ", Sum (CASE WHEN DETL_POSTING_AMT < 0 THEN -1 * DETL_POSTING_AMT ELSE 0 END) ACCT_CR" _
            & ", Count (*) ACCT_TRANS" _
            & " from GLTDETL1,GLTJRNL1,GLTACCT1 " _
            & " where GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE " _
            & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" _
            & "   and GLTDETL1.OPS_YYYYPP = :PARM1" _
            & " group by GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE" _
            & ", GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" _
            & ", GLTACCT1.ACCT_DESC"
            Create_TDA(.Tables.Add, "GLTSUMJ2", "**", 0, False, "V", 5)
            .Tables("GLTSUMJ2").Columns("ACCT_TRANS").DataType = GetType(System.Int32)

            .Relations.Add("GLTSUMJ2", _
            .Tables("GLTSUMJ1").Columns("JOURNAL_TYPE"), _
            .Tables("GLTSUMJ2").Columns("JOURNAL_TYPE"))

            .Tables("GLTSUMJ1").Columns("ACCT_DR").Expression = "SUM(Child(GLTSUMJ2).ACCT_DR)"
            .Tables("GLTSUMJ1").Columns("ACCT_CR").Expression = "SUM(Child(GLTSUMJ2).ACCT_CR)"
            .Tables("GLTSUMJ1").Columns("ACCT_TRANS").Expression = "SUM(Child(GLTSUMJ2).ACCT_TRANS)"
        End With

        grdGLTTBAL1.DataSource = dst.Tables("GLTTBAL0")
        grdGLTSUMJ1.DataSource = dst.Tables("GLTSUMJ1")


        Call Add_ACCT_TYPEs("GLTTBAL0", True)

        Call Get_PARM("GLTPARM1")

        Call Set_SEGS(grdGLTTBAL1, "GLTTBAL1")
        Call Set_SEGS(grdGLTSUMJ1, "GLTSUMJ2")

        grdGLTSUMJ1.DisplayLayout.Bands("GLTSUMJ2").SummaryFooterCaption = "Totals for Journal Type: [JOURNAL_TYPE] [JOURNAL_TYPE_DESC]"

        Create_Lookup("GLTACCT1")


        Call Create_Summary(grdGLTTBAL1, "ACCT_BEG_BAL")
        Call Create_Summary(grdGLTTBAL1, "ACCT_DR")
        Call Create_Summary(grdGLTTBAL1, "ACCT_CR")
        Call Create_Summary(grdGLTTBAL1, "ACCT_END_BAL")
        Call Create_Summary(grdGLTTBAL1, "ACCT_TRANS")

        Call Create_Summary(grdGLTTBAL1, "ACCT_BEG_BAL", , "GLTTBAL1")
        Call Create_Summary(grdGLTTBAL1, "ACCT_DR", , "GLTTBAL1")
        Call Create_Summary(grdGLTTBAL1, "ACCT_CR", , "GLTTBAL1")
        Call Create_Summary(grdGLTTBAL1, "ACCT_END_BAL", , "GLTTBAL1")
        Call Create_Summary(grdGLTTBAL1, "ACCT_TRANS", , "GLTTBAL1")
        Call Create_Summary(grdGLTTBAL1, "ACCT_CODE", "Count", "GLTTBAL1")


        Call Create_Summary(grdGLTSUMJ1, "ACCT_DR")
        Call Create_Summary(grdGLTSUMJ1, "ACCT_CR")
        Call Create_Summary(grdGLTSUMJ1, "ACCT_TRANS")

        Call Create_Summary(grdGLTSUMJ1, "ACCT_DR", , "GLTSUMJ2")
        Call Create_Summary(grdGLTSUMJ1, "ACCT_CR", , "GLTSUMJ2")
        Call Create_Summary(grdGLTSUMJ1, "ACCT_TRANS", , "GLTSUMJ2")
        Call Create_Summary(grdGLTSUMJ1, "ACCT_CODE", "Count", "GLTSUMJ2")


        With grdGLTTBAL1.DisplayLayout
            .Bands("GLTTBAL0").SortedColumns.Clear()
            .Bands("GLTTBAL0").SortedColumns.Add("ACCT_TYPE_SEQ", False)
            .Bands("GLTTBAL1").SortedColumns.Clear()
            .Bands("GLTTBAL1").SortedColumns.Add("ACCT_CODE", False)
            .Bands("GLTTBAL1").SortedColumns.Add("SEG2_CODE", False)
            .Bands("GLTTBAL1").SortedColumns.Add("SEG3_CODE", False)
            .Bands("GLTTBAL1").SortedColumns.Add("SEG4_CODE", False)
        End With

        With grdGLTSUMJ1.DisplayLayout
            '.Bands("GLTSUMJ1").SortedColumns.Clear()
            '.Bands("GLTSUMJ1").SortedColumns.Add("ACCT_TYPE_SEQ", False)
            .Bands("GLTSUMJ2").SortedColumns.Clear()
            .Bands("GLTSUMJ2").SortedColumns.Add("ACCT_CODE", False)
            .Bands("GLTSUMJ2").SortedColumns.Add("SEG2_CODE", False)
            .Bands("GLTSUMJ2").SortedColumns.Add("SEG3_CODE", False)
            .Bands("GLTSUMJ2").SortedColumns.Add("SEG4_CODE", False)
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

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
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
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

        dst.EnforceConstraints = False
        dst.Tables("GLTTBAL1").Rows.Clear()
        dst.Tables("GLTSUMJ1").Rows.Clear()
        dst.Tables("GLTSUMJ2").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("OPS_YYYYPP").Text = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Account Summary Data")

        grdGLTTBAL1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        Call Save_Header_Fields(UltraGroupBox1)

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
        Call ASCMAIN1.AnalyzeTable(GLTACCT3x)

        sql = "Select GLTACCT3.ACCT_CODE" _
            & ", GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
            & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" _
            & ", GLTACCT3.ACCT_BEG_BAL" & sql_BEG_BAL & " ACCT_BEG_BAL" _
            & ", GLTACCT3.ACCT_BEG_BAL" & sql_END_BAL & " ACCT_END_BAL" _
            & ", GLTACCT3X.ACCT_DR" _
            & ", GLTACCT3X.ACCT_CR" _
            & ", GLTACCT3X.ACCT_TRANS" _
            & " from " & GLTACCT3 & " GLTACCT3,GLTACCT1," & GLTACCT3x & " GLTACCT3X" _
            & " where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE " _
            & "   and GLTACCT3.ACCT_YEAR = '" & ACCT_YEAR & "'" _
            & " and GLTACCT3.ACCT_CODE = GLTACCT3X.ACCT_CODE (+)" _
            & " and GLTACCT3.SEG2_CODE = GLTACCT3X.SEG2_CODE (+)" _
            & " and GLTACCT3.SEG3_CODE = GLTACCT3X.SEG3_CODE (+)" _
            & " and GLTACCT3.SEG4_CODE = GLTACCT3X.SEG4_CODE (+)"
        Set_SelectCommand("GLTTBAL1", sql)
        Call Fill_Records("GLTTBAL1")

        For Each row As DataRow In dst.Tables("GLTTBAL0").Rows
            Dim ACCT_TYPE As String = row.Item("ACCT_TYPE")
            Dim sqlx As String = "ACCT_TYPE = '" & ACCT_TYPE & "'"
            With dst.Tables("GLTTBAL1")
                For Each COLUMN_NAME In New String() {"ACCT_BEG_BAL", "ACCT_END_BAL", "ACCT_DR", "ACCT_CR", "ACCT_TRANS"}
                    row.Item(COLUMN_NAME) = Val(.Compute("SUM (" & COLUMN_NAME & ")", sqlx) & "")
                Next
            End With
        Next

        Call Fill_Records("GLTSUMJ1", HFs("OPS_YYYYPP"))
        Call Fill_Records("GLTSUMJ2", HFs("OPS_YYYYPP"))

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()


        Call CommitTrans("Update Complete")
    End Sub
#End Region

End Class