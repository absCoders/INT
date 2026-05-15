Public Class ARFCOLL1

    Dim RYP As String
    Dim FYP As String
    Dim ARTCOLL1 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_ARTCOLL1("", "")
            ASCMAIN1.sql = "Select * from " & ARTCOLL1
            Create_TDA(.Tables.Add, "ARTCOLL1", "**", 0, False, "", 0)
            .Tables("ARTCOLL1").Columns.Add("AR_GROWTH", GetType(System.Decimal), "ISNULL(AR,0) - ISNULL(BEG,0)")
        End With

        grdARTCOLL1.DataSource = dst.Tables("ARTCOLL1")

        Call Create_Summary(grdARTCOLL1, "CUST_CODE", "Count")
        Call Create_Summary(grdARTCOLL1, "BEG")
        Call Create_Summary(grdARTCOLL1, "AGE_1")
        Call Create_Summary(grdARTCOLL1, "AGE_2")
        Call Create_Summary(grdARTCOLL1, "AGE_3")
        Call Create_Summary(grdARTCOLL1, "AGE_4")
        Call Create_Summary(grdARTCOLL1, "AR")
        Call Create_Summary(grdARTCOLL1, "SALES")
        Call Create_Summary(grdARTCOLL1, "PYMTS")
        Call Create_Summary(grdARTCOLL1, "AR_GROWTH")

        grdARTCOLL1.DisplayLayout.UseFixedHeaders = True
        With grdARTCOLL1.DisplayLayout.Bands("ARTCOLL1")
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("SREP_CODE").Header.Fixed = True
        End With

        With grdARTCOLL1.DisplayLayout.Bands("ARTCOLL1")
            .Columns("BEG").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("AGE_1").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("AGE_2").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("AGE_3").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("AGE_4").CellAppearance.BackColor = Drawing.Color.LightYellow
            .Columns("AR").CellAppearance.BackColor = Drawing.Color.LightBlue
            '.Columns("SALES").CellAppearance.BackColor = Drawing.Color.LightGreen
            '.Columns("PYMTS").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("AR_GROWTH").CellAppearance.BackColor = Drawing.Color.LightGreen
        End With

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Call Validate_Code("OPS_YYYYPP")
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

            Case "Print"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdARTCOLL1.Visible = tf

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ARTCOLL1").Rows.Clear()
        dst.EnforceConstraints = True

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Data")

        Call Save_Header_Fields(UltraGroupBox1)

        Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
        RYP = z
        FYP = ASCMAIN1.Period_Calc(RYP, 1)

        Create_ARTCOLL1(FYP, RYP)

        Call ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)
        Fill_Records("ARTCOLL1")
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdARTCOLL1, "SSB", "Show Filter", "Show GroupBox", "Customer Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case "grdARTCOLL1"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool
                Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden

            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select


    End Sub

#End Region

    Sub Create_ARTCOLL1(ByVal FYP As String, ByVal RYP As String)

        ASCMAIN1.sql = "SELECT X.*, ARTCUST6.CUST_LAST_PMT_DATE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE FROM " _
        & " (SELECT CUST_CODE, SUM (BEG) BEG" _
        & ", SUM (AGE_1) AGE_1, SUM (AGE_2) AGE_2, SUM (AGE_3) AGE_3, SUM (AGE_4) AGE_4" _
        & ", SUM (AR) AR, SUM (SALES) SALES, SUM (PYMTS) PYMTS" _
        & " FROM (" _
        & " SELECT CUST_CODE, 0 BEG, 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4" _
        & ", SUM (INV_BALANCE) AR, 0 SALES, 0 PYMTS FROM ARTOPEN1 GROUP BY CUST_CODE" _
        & " UNION" _
        & " SELECT CUST_CODE, TOTAL_DUE BEG, AGE_1, AGE_2, AGE_3, AGE_4, " _
        & "0 AR, 0 SALES, 0 PYMTS FROM ARTSTMT1 WHERE OPS_YYYYPP = '" & RYP & "'" _
        & " UNION" _
        & " SELECT GROUP_NO CUST_CODE, 0 BEG" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4" _
        & ", 0 AR, SUM (INV_TOTAL_AMT) SALES, 0 PYMTS" _
        & " FROM BATINVH1 WHERE OPS_YYYYPP >= '" & FYP & "' GROUP BY GROUP_NO" _
        & " UNION " _
        & " SELECT CUST_CODE, 0 BEG" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4" _
        & ", 0 AR, 0 SALES, SUM (CUST_PYMT_AMT) PYMTS" _
        & " FROM ARTPYMT1,ARTPYMT2 WHERE ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " AND ARTPYMT2.CUST_PYMT_AMT <> 0 AND ARTPYMT1.OPS_YYYYPP >= '" & FYP & "' GROUP BY CUST_CODE" _
        & ") GROUP BY CUST_CODE) X, ARTCUST6, ARTCUST1" _
        & " WHERE ARTCUST6.CUST_CODE (+) = X.CUST_CODE AND ARTCUST1.CUST_CODE = X.CUST_CODE"
        If ARTCOLL1 = "" Then
            ARTCOLL1 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCOLL1)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCOLL1 & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Call Print_Report_Begin()

        Call Print_Report_End()
    End Sub
End Class