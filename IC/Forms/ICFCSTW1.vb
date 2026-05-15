Imports Infragistics.Win.UltraWinGrid

Public Class ICFCSTW1



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = ICCMAIN1.Get_sqlICTCSTW1
            Create_TDA(.Tables.Add, "ICTCSTW1", "**", 0, False, "VVV")
            With .Tables("ICTCSTW1").Columns
                Dim DIFF As String = "ISNULL(CALC_VCOST,0) - ISNULL(VCOST,0)"

                ' the above isn an ado.net expression - column with this expression is being added below
                .Add("VCOST_DIFF", GetType(System.Decimal), DIFF)
            End With
            ASCMAIN1.sql = "select h.*, i.item_desc " &
        "from ictvcsth h " &
        "join ictitem1 i on h.item_code = i.item_code " &
        "where h.ops_yyyypp = :PARM1" &
        " and h.vend_code = :PARM2"
            Create_TDA(.Tables.Add, "ICTCSTH1", "**", 0, False, "VV")
            With dst.Tables("ICTCSTH1").Columns
                Dim DIFF As String = "ISNULL(CALC_VCOST,0) - ISNULL(VCOST,0)"
                .Add("VCOST_DIFF", GetType(System.Decimal), DIFF)
            End With
        End With

        grdICTCSTW1.DataSource = dst.Tables("ICTCSTW1")
        Create_Summary(grdICTCSTW1, "ITEM_CODE", "Count")

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
                EntryMode = "I"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

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

        'UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ICTCSTW1").Rows.Clear()
        dst.Tables("ICTCSTH1").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
        Absx1.txtFor("VEND_CODE").Text = "IPSA"
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Item Cost Data")

        'grdGLTTBAL1.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.BottomFixed

        'Call Save_Header_Fields(UltraGroupBox1)

        'Dim ACCT_YEAR As String = Mid$(HFs("OPS_YYYYPP"), 1, 4)
        'Dim P As Integer = Val(Mid$(HFs("OPS_YYYYPP"), 5, 2))

        'Dim sql_BEG_BAL As String = ""
        'Dim sql_END_BAL As String = ""
        'For i As Integer = 1 To P
        '    Dim z As String = " + NVL(GLTACCT3.ACCT_ACT_P" & Format(i, "00") & ",0)"
        '    If i < P Then
        '        sql_BEG_BAL = sql_BEG_BAL & z
        '    End If
        '    sql_END_BAL = sql_END_BAL & z
        'Next

        'Dim GLTACCT3 As String = GL_Prep(ACCT_YEAR, ACCT_YEAR)

        'Dim sql As String = ""


        'sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE " _
        '    & ", Sum (CASE WHEN DETL_POSTING_AMT > 0 THEN DETL_POSTING_AMT ELSE 0 END) ACCT_DR" _
        '    & ", Sum (CASE WHEN DETL_POSTING_AMT < 0 THEN -1 * DETL_POSTING_AMT ELSE 0 END) ACCT_CR" _
        '    & ", Count (*) ACCT_TRANS" _
        '    & " from GLTDETL1 where OPS_YYYYPP = '" & HFs("OPS_YYYYPP") & "'" _
        '    & " group by ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"
        'Dim GLTACCT3x As String = ASCMAIN1.Temp_Table(sql)
        'ASCDATA1.ExecuteSQL("Alter Table " & GLTACCT3x & " Add Primary Key (ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE)")
        'Call ASCMAIN1.AnalyzeTable(GLTACCT3x)

        'sql = "Select GLTACCT3.ACCT_CODE" _
        '    & ", GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
        '    & ", GLTACCT1.ACCT_TYPE, GLTACCT1.ACCT_DESC" _
        '    & ", GLTACCT3.ACCT_BEG_BAL" & sql_BEG_BAL & " ACCT_BEG_BAL" _
        '    & ", GLTACCT3.ACCT_BEG_BAL" & sql_END_BAL & " ACCT_END_BAL" _
        '    & ", GLTACCT3X.ACCT_DR" _
        '    & ", GLTACCT3X.ACCT_CR" _
        '    & ", GLTACCT3X.ACCT_TRANS" _
        '    & " from " & GLTACCT3 & " GLTACCT3,GLTACCT1," & GLTACCT3x & " GLTACCT3X" _
        '    & " where GLTACCT1.ACCT_CODE (+)= GLTACCT3.ACCT_CODE " _
        '    & "   and GLTACCT3.ACCT_YEAR = '" & ACCT_YEAR & "'" _
        '    & " and GLTACCT3.ACCT_CODE = GLTACCT3X.ACCT_CODE (+)" _
        '    & " and GLTACCT3.SEG2_CODE = GLTACCT3X.SEG2_CODE (+)" _
        '    & " and GLTACCT3.SEG3_CODE = GLTACCT3X.SEG3_CODE (+)" _
        '    & " and GLTACCT3.SEG4_CODE = GLTACCT3X.SEG4_CODE (+)"
        'Set_SelectCommand("GLTTBAL1", sql)
        'Call Fill_Records("GLTTBAL1")

        'For Each row As DataRow In dst.Tables("GLTTBAL0").Rows
        '    Dim ACCT_TYPE As String = row.Item("ACCT_TYPE")
        '    Dim sqlx As String = "ACCT_TYPE = '" & ACCT_TYPE & "'"
        '    With dst.Tables("GLTTBAL1")
        '        For Each COLUMN_NAME In New String() {"ACCT_BEG_BAL", "ACCT_END_BAL", "ACCT_DR", "ACCT_CR", "ACCT_TRANS"}
        '            row.Item(COLUMN_NAME) = Val(.Compute("SUM (" & COLUMN_NAME & ")", sqlx) & "")
        '        Next
        '    End With
        'Next

        'Call Fill_Records("GLTSUMJ1", HFs("OPS_YYYYPP"))
        Dim opsYYYYPP As String = If(String.IsNullOrWhiteSpace(Absx1.txtFor("OPS_YYYYPP")?.Text), ASCMAIN1.CYP, Absx1.txtFor("OPS_YYYYPP").Text)
        Dim vendCode As String = If(String.IsNullOrWhiteSpace(Absx1.txtFor("VEND_CODE")?.Text), "IPSA", Absx1.txtFor("VEND_CODE").Text)

        Dim isCurrentPeriod As Boolean = (opsYYYYPP = ASCMAIN1.CYP)

        If isCurrentPeriod Then
            ASCMAIN1.sql = ICCMAIN1.Get_sqlICTCSTW1
            Call Fill_Records("ICTCSTW1", New Object() {opsYYYYPP, opsYYYYPP, vendCode})
            grdICTCSTW1.DataSource = dst.Tables("ICTCSTW1")
            grdICTCSTW1.Refresh()
            grdICTCSTW1.Text = $"Saleable items for {vendCode} with OH/PO quantities and/or activity this period"
        Else
            Dim sqlHist As String =
        "select h.*, i.item_desc " &
        "from ictvcsth h " &
        "join ictitem1 i on h.item_code = i.item_code " &
        "where h.ops_yyyypp = :PARM1" &
        " and h.vend_code = :PARM2"

            dst.EnforceConstraints = False
            dst.Tables("ICTCSTH1").Clear()
            dst.EnforceConstraints = True

            Call Fill_Records("ICTCSTH1", New Object() {opsYYYYPP, vendCode})

            grdICTCSTW1.DataSource = dst.Tables("ICTCSTH1")
            grdICTCSTW1.Refresh()

            grdICTCSTW1.Text = $"Item costs (history) for period {opsYYYYPP}"
        End If


        'Call Fill_Records("ICTCSTW1", New Object() {opsYYYYPP, opsYYYYPP, vendCode})
        'grdICTCSTW1.Text = $"Saleable items for {vendCode} with OH/PO quantities and/or activity this period"
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'ASCDATA1.ExecuteSQL(sql, "VNNVVV", New Object() _
        '        {ASCMAIN1.SESSION_NO, selectionNo, RE_XNO _
        '        , ASCMAIN1.USER_ID, errMsg, usefulTrace})
        ''only execute if oracle is available

        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCSTW1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If
        If e.SourceControl Is Nothing OrElse Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
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

            'Select Case e.SourceControl.Name

            '    'Case "grdSOTORDRX"
            '    '    tlb_btn = DirectCast(tlb.Tools("Multi-Order Edit"), UltraWinToolbars.ButtonTool)
            '    '    tlb_btn.SharedProps.Visible = False

            '    '    If (MENU_ITEM_OBJECT = "SOFORDR1" And grdSOTORDRX.Selected.Rows.Count > 1) Then
            '    '        tlb_btn.SharedProps.Visible = True
            '    '    End If

            'End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

    Private Sub grdICTCSTW1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTCSTW1.InitializeRow

        If e.Row.Band Is Nothing Then Exit Sub

        If e.Row.Band.Key <> "ICTCSTW1" AndAlso e.Row.Band.Key <> "ICTCSTH1" Then Exit Sub

        If Not e.Row.Cells.Exists("VCOST_DIFF") Then Exit Sub

        Dim diff As Decimal = 0D
        If e.Row.Cells("VCOST_DIFF").Value IsNot Nothing AndAlso IsNumeric(e.Row.Cells("VCOST_DIFF").Value) Then
            diff = Convert.ToDecimal(e.Row.Cells("VCOST_DIFF").Value)
        End If

        If diff <> 0D Then
            e.Row.Cells("VCOST_DIFF").Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("VCOST_DIFF").Appearance.BackColor = Drawing.Color.Empty
        End If

    End Sub

#End Region

End Class