Imports System.Math

Public Class GLFPEND1
    ' Calculate next period NYP from Current Period RYP
    Dim RYP As String
    Dim NYP As String
    Dim LYP As String
    Dim ACCT_END_BAL_TOTAL As Double

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            .Tables.Add("GLTPEND0")
            .Tables("GLTPEND0").Columns.Add("ACCT_TYPE", GetType(System.String))
            .Tables("GLTPEND0").Columns.Add("ACCT_TYPE_DESC", GetType(System.String))
            .Tables("GLTPEND0").Columns.Add("ACCT_END_BAL", GetType(System.Decimal))
            .Tables("GLTPEND0").PrimaryKey = New DataColumn() {.Tables("GLTPEND0").Columns("ACCT_TYPE")}

            ASCMAIN1.sql = "Select GLTACCT1.ACCT_TYPE, GLTACCT3.ACCT_CODE, GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE, GLTACCT1.ACCT_DESC, ACCT_BEG_BAL ACCT_END_BAL from GLTACCT3,GLTACCT1 where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE and GLTACCT3.ACCT_YEAR = :PARM1"
            '            Create_TDA(.Tables.Add, "GLTPEND1", "**", 0, True, False, "V", 5)
            .Tables.Add(Create_ResultSet("GLTPEND1", "V", 5))

            .Relations.Add("GLTPEND1", _
            .Tables("GLTPEND0").Columns("ACCT_TYPE"), _
            .Tables("GLTPEND1").Columns("ACCT_TYPE"))

            Create_TDA(.Tables.Add, "GLTACCT3", "*")
            Create_TDA(.Tables.Add, "GLTPARM1", "*")
        End With

        Call Add_ACCT_TYPEs("GLTPEND0")
        Call Get_PARM("GLTPARM1")
        Call Set_SEGS(grdGLTPEND1, "GLTPEND1")
        Create_Lookup("GLTACCT1")

        ASCMAIN1.sql = "Select Sum (ACCT_BEG_BAL)"
        For i As Integer = 1 To 13
            ASCMAIN1.sql &= ", Sum (ACCT_ACT_P" & Format$(i, "00") & ")"
        Next i
        ASCMAIN1.sql &= " from GLTACCT3 where ACCT_YEAR = :PARM1"
        Call Create_ResultSet("TEST_BALANCE", "V")

        grdGLTPEND1.DataSource = dst.Tables("GLTPEND0")

        Call Create_Summary(grdGLTPEND1, "ACCT_END_BAL")
        Call Create_Summary(grdGLTPEND1, "ACCT_END_BAL", , "GLTPEND1")
        Call Create_Summary(grdGLTPEND1, "ACCT_CODE", "Count", "GLTPEND1")

        If ASCMAIN1.CLIENT = "INT" Then
            optAction.Visible = False
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Dim periodToClose As String = LYP

                ASCMAIN1.sql = $"Select COUNT(*) from GLTJRNL1 where JOURNAL_APPR_STATUS = 'P' and OPS_YYYYPP <= {RYP}"

                Dim pendingCnt As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql)

                If pendingCnt > 0 Then
                    EMsg &= vbCr & $"There are {pendingCnt} pending journal entries dated in or before period " & periodToClose & "." & vbCr & vbCr & "All journals must be approved or rejected before closing the period."
                    Exit Select
                End If

                If optAction.Value = "R" Then
                    If MsgBox("OK to Proceed?", _
                              MsgBoxStyle.YesNo, _
                              "You are Re-Opening a Closed GL Period") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                If Round(ACCT_END_BAL_TOTAL, 2) <> 0 Then
                    EMsg &= vbCr & "GL is Out of Balance"
                End If

                ' Check that all periods are in balance in the current year
                Call Fill_Record("TEST_BALANCE", Mid(RYP, 1, 4))

                For i As Integer = 0 To 13
                    If Abs(Sign(Val(cdr.Item(i) & ""))) <> 0 Then
                        EMsg &= "Out of Balance in Period " & CStr(i) & " of " & Mid$(RYP, 1, 4)
                    End If
                Next i

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

                If optAction.Value = "R" Then

                    BeginTrans()

                    TAC.TACMAIN1.Record_Event("GLTPARM1", "Z", DATETIME_STAMP, ASCMAIN1.USER_ID, "REOPEN", LYP & " Re-Opened")
                    Fill_Records("GLTPARM1", "Z")
                    Dim GL_PARM_CURRENT_YYYYPP As String = TBLs("GLTPARM1").Rows(0).Item("GL_PARM_CURRENT_YYYYPP") & ""

                    TBLs("GLTPARM1").Rows(0).Item("GL_PARM_CURRENT_YYYYPP") = LYP
                    Update_Record_TDA("GLTPARM1")

                    If Mid(LYP, 5, 2) = "12" Then
                        ASCMAIN1.sql = "Update GLTACCT3 Set ACCT_BEG_BAL = 0 where ACCT_YEAR = '" & Mid(GL_PARM_CURRENT_YYYYPP, 1, 4) & "'"
                        ASCDATA1.ExecuteSQL()
                    End If

                    CommitTrans()

                    MsgBox("Period " & LYP & " has been Re-Opened", MsgBoxStyle.OkOnly, "Success")

                    Call Mode_Settings(False)

                Else
                    EntryMode = "E"
                    Call Load_Record()
                    Call Mode_Settings(True)
                End If


            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdGLTPEND1.Visible = tf
        optAction.Value = "C"

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        Call Get_PARM("GLTPARM1")
        dst.Tables("GLTPEND1").Rows.Clear()

        RYP = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
        NYP = ASCMAIN1.Period_Calc(RYP, 1)
        LYP = ASCMAIN1.Period_Calc(RYP, -1)

        Set_Period()

    End Sub

    Sub Set_Period()
        If optAction.Value = "C" Then
            Absx1.txtFor("OPS_YYYYPP").Text = RYP
            UltraExplorerBar1.Groups("Screen Control").Items("Load").Text = "Load"
        Else
            Absx1.txtFor("OPS_YYYYPP").Text = LYP
            UltraExplorerBar1.Groups("Screen Control").Items("Load").Text = "Re-Open Period"
        End If
    End Sub
    Sub Load_Record()

        'grdGLTPEND1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        ASCMAIN1.Progress("Now Loading GL Activity Summary")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Save_Header_Fields(UltraGroupBox1)
        Dim ACCT_YEAR As String = Mid$(HFs("OPS_YYYYPP"), 1, 4)

        EnforceConstraints(False)

        Fill_Records("GLTPARM1", "Z")
        Fill_Records("GLTPEND1", ACCT_YEAR)

        For Each row As DataRow In dst.Tables("GLTPEND1").Select("")
            Dim rowGLTACCT3 As DataRow = LookUp("GLTACCT3", New String() _
                {row.Item("ACCT_CODE"), _
                 row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), _
                 ACCT_YEAR})
            Dim T As Decimal = Val(rowGLTACCT3.Item("ACCT_BEG_BAL") & "")
            For P As Integer = 1 To Val(Mid(HFs("OPS_YYYYPP"), 5, 2))
                T += Val(rowGLTACCT3.Item("ACCT_ACT_P" & Format(P, "00")) & "")
            Next
            row.Item("ACCT_END_BAL") = T
        Next

        ACCT_END_BAL_TOTAL = 0
        For Each row As DataRow In dst.Tables("GLTPEND0").Rows
            Dim ACCT_TYPE As String = row.Item("ACCT_TYPE")
            Dim ACCT_END_BAL As Decimal = _
            Val(dst.Tables("GLTPEND1").Compute("SUM (ACCT_END_BAL)", _
                "ACCT_TYPE = '" & ACCT_TYPE & "'") & "")
            row.Item("ACCT_END_BAL") = ACCT_END_BAL
            '            ACCT_END_BAL_TOTAL = ACCT_END_BAL_TOTAL + ACCT_END_BAL
            ACCT_END_BAL_TOTAL += ACCT_END_BAL
        Next
        EnforceConstraints(True)

        grdGLTPEND1.DisplayLayout.Bands("GLTPEND1").SortedColumns.Clear()
        grdGLTPEND1.DisplayLayout.Bands("GLTPEND1").SortedColumns.Add("ACCT_CODE", False)
        grdGLTPEND1.DisplayLayout.Bands("GLTPEND1").SortedColumns.Add("SEG2_CODE", False)
        grdGLTPEND1.DisplayLayout.Bands("GLTPEND1").SortedColumns.Add("SEG3_CODE", False)
        grdGLTPEND1.DisplayLayout.Bands("GLTPEND1").SortedColumns.Add("SEG4_CODE", False)

        Call ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
        Application.DoEvents()

    End Sub

    Sub Update_Record()
        Dim Sql As String
        Dim TT As String = ""
        If Mid$(RYP, 1, 4) <> Mid$(NYP, 1, 4) Then
            TT = GL_Prep(Mid$(NYP, 1, 4), Mid$(NYP, 1, 4))
        End If

        BeginTrans()

        If Mid$(RYP, 1, 4) <> Mid$(NYP, 1, 4) Then
            ' THIS NEXT SQL DOES NOT APPEAR TO BE NEC - IF THE BEG BAL CONTAINS ANYTHING AT THIS POINT, THEN THE TT TABLE IS WHACKED
            ' THERE SHOULD BE NOTHING IN THE BEG BAL FOR FUTURE YEARS - NO ABSOLUTION ROUTINES WOULD HAVE PUT A BEG BAL IN A FUTURE YEAR

            Sql = "Update GLTACCT3 set ACCT_BEG_BAL = 0 where ACCT_YEAR = '" & Mid$(NYP, 1, 4) & "'"
            ASCDATA1.ExecuteSQL(Sql)

            dst.Tables("GLTACCT3").Rows.Clear()

            Sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_BEG_BAL from " & TT
            Sql = Sql & " where ACCT_YEAR = '" & Mid$(NYP, 1, 4) & "'"
            Sql = Sql & "   and ACCT_BEG_BAL <> 0"
            For Each row As DataRow In ASCDATA1.GetDataTable(Sql, TT).Rows

                Fill_Record("GLTACCT3", New String() _
                {row.Item("ACCT_CODE"), _
                row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), _
                Mid$(NYP, 1, 4)}, True)
                Dim rowGLTACCT3 As DataRow = dst.Tables("GLTACCT3").Rows.Find(New Object() {row.Item("ACCT_CODE"), _
                row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), _
                Mid$(NYP, 1, 4)})
                rowGLTACCT3.Item("ACCT_BEG_BAL") = row.Item("ACCT_BEG_BAL")
                Update_Record_TDA("GLTACCT3")
            Next

            Call Fill_Record("TEST_BALANCE", Mid(NYP, 1, 4))
            For i As Integer = 0 To 13
                If Abs(Sign(Val(cdr.Item(i) & ""))) <> 0 Then
                    Rollback("Out of Balance in Period " & CStr(i) & " of " & Mid$(NYP, 1, 4))
                    Exit Sub
                End If
            Next i
        End If

        ' Update Period Control Record
        ' (Note that this is the only thing that happens when closing a period except for at year-end)
        'dst.Tables("GLTPARM1").Rows(0).Item("GL_PARM_CURRENT_YYYYPP") = NYP
        Dim GL_PARM_CURRENT_YYYYPP_prev As String = TBLs("GLTPARM1").Rows(0).Item("GL_PARM_CURRENT_YYYYPP")
        TBLs("GLTPARM1").Rows(0).Item("GL_PARM_CURRENT_YYYYPP") = NYP

        If TBLs("GLTPARM1").Rows(0).Item("GL_PARM_SOFT_CLOSE_YYYYPP") < GL_PARM_CURRENT_YYYYPP_prev Then
            TBLs("GLTPARM1").Rows(0).Item("GL_PARM_SOFT_CLOSE_YYYYPP") = GL_PARM_CURRENT_YYYYPP_prev
        End If
        Update_Record_TDA("GLTPARM1")

        CommitTrans("Update Complete")
    End Sub
#End Region

    Private Sub optAction_ValueChanged(sender As Object, e As EventArgs) Handles optAction.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        If optAction.Value = "C" Then
            lblPeriodToClose.Text = "Period to Close"
        Else
            lblPeriodToClose.Text = "Period to Re-Open"
        End If

        Set_Period()
    End Sub
End Class