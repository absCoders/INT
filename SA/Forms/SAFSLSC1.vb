Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles

Public Class SAFSLSC1

    Dim RYP As String
    Dim FYP As String
    Dim SATSLSC0 As String
    Dim SLS_MTD_TOTAL As Double
    Dim SLS_YTD_TOTAL As Double


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Create_SATSLSC0("", "")
            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, CUST_CLASS_CODE, MARGIN from " & SATSLSC0
            Create_TDA(.Tables.Add, "SATSLSC1", "**", 0, False, "", 0)
            For i As Integer = 1 To 12
                .Tables("SATSLSC1").Columns.Add("GROSS_" & Format(i, "00"), GetType(System.Int32))
                .Tables("SATSLSC1").Columns.Add("RET_" & Format(i, "00"), GetType(System.Int32))
                .Tables("SATSLSC1").Columns.Add("NET_" & Format(i, "00"), GetType(System.Int32))
                .Tables("SATSLSC1").Columns.Add("RETL_" & Format(i, "00"), GetType(System.Int32))
            Next
            .Tables("SATSLSC1").Columns.Add("GROSS_TOT", GetType(System.Int32))
            .Tables("SATSLSC1").Columns.Add("RET_TOT", GetType(System.Int32))
            .Tables("SATSLSC1").Columns.Add("NET_TOT", GetType(System.Int32))
            .Tables("SATSLSC1").Columns.Add("RETL_TOT", GetType(System.Int32))
        End With

        grdSATSLSC1.DataSource = dst.Tables("SATSLSC1")
        With grdSATSLSC1.DisplayLayout.Bands(0)
            .Groups.Add("CODES", "Customer")
            .Columns("CUST_CODE").Group = .Groups("CODES")
            .Columns("CUST_NAME").Group = .Groups("CODES")
            .Columns("CUST_CLASS_CODE").Group = .Groups("CODES")
            .Columns("MARGIN").Group = .Groups("CODES")
            .Columns("CUST_CODE").Header.Caption = "Cust"
            .Columns("CUST_NAME").Header.Caption = "Name"
            .Columns("CUST_CLASS_CODE").Header.Caption = "Class"
            .Columns("MARGIN").Header.Caption = "Margin"
            .Columns("CUST_CODE").Width = 100
            .Columns("CUST_NAME").Width = 300
            .Columns("CUST_CLASS_CODE").Width = 60
            .Columns("MARGIN").Width = 60
            .Groups.Add("Total")
            .Columns("GROSS_TOT").Group = .Groups("Total")
            .Columns("RET_TOT").Group = .Groups("Total")
            .Columns("NET_TOT").Group = .Groups("Total")
            .Columns("RETL_TOT").Group = .Groups("Total")
            .Columns("GROSS_TOT").Header.Caption = "Gross"
            .Columns("RET_TOT").Header.Caption = "Returns"
            .Columns("NET_TOT").Header.Caption = "Net"
            .Columns("RETL_TOT").Header.Caption = "Net@Retl"
            .Groups("Total").Header.Appearance.BackColor = Drawing.Color.Cyan
            .Columns("GROSS_TOT").Header.Appearance.BackColor = Drawing.Color.Cyan
            .Columns("RET_TOT").Header.Appearance.BackColor = Drawing.Color.Cyan
            .Columns("NET_TOT").Header.Appearance.BackColor = Drawing.Color.Cyan
            .Columns("RETL_TOT").Header.Appearance.BackColor = Drawing.Color.Cyan
            .Columns("GROSS_TOT").Width = 100
            .Columns("RET_TOT").Width = 100
            .Columns("NET_TOT").Width = 100
            .Columns("RETL_TOT").Width = 100
            .Groups("Total").Header.Appearance.TextHAlign = HAlign.Center
            For i As Integer = 1 To 12
                .Groups.Add("M_" & Format(i, "00"))
                .Columns("GROSS_" & Format(i, "00")).Group = .Groups("M_" & Format(i, "00"))
                .Columns("RET_" & Format(i, "00")).Group = .Groups("M_" & Format(i, "00"))
                .Columns("NET_" & Format(i, "00")).Group = .Groups("M_" & Format(i, "00"))
                .Columns("RETL_" & Format(i, "00")).Group = .Groups("M_" & Format(i, "00"))
                .Columns("GROSS_" & Format(i, "00")).Header.Caption = "Gross"
                .Columns("RET_" & Format(i, "00")).Header.Caption = "Returns"
                .Columns("NET_" & Format(i, "00")).Header.Caption = "Net"
                .Columns("RETL_" & Format(i, "00")).Header.Caption = "Net@Retl"
                .Columns("GROSS_" & Format(i, "00")).Width = 90
                .Columns("RET_" & Format(i, "00")).Width = 90
                .Columns("NET_" & Format(i, "00")).Width = 90
                .Columns("RETL_" & Format(i, "00")).Width = 90
                If i Mod 2 = 1 Then
                    .Columns("GROSS_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Yellow
                    .Columns("RET_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Yellow
                    .Columns("NET_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Yellow
                    .Columns("RETL_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Yellow
                Else
                    .Columns("GROSS_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Cyan
                    .Columns("RET_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Cyan
                    .Columns("NET_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Cyan
                    .Columns("RETL_" & Format(i, "00")).Header.Appearance.BackColor = Drawing.Color.Cyan
                End If
            Next

            End With

            'For Each T As String In New String() {"MTD", "YTD"}
            '    Dim C As System.Drawing.Color = Drawing.Color.Yellow
            '    Dim W As Int32 = 80
            '    If T = "YTD" Then
            '        C = Drawing.Color.LightGreen
            '        W = 90
            '    End If

            '    .Columns("SLS_" & T).Header.Appearance.BackColor = C
            '    .Columns("CGS_" & T).Header.Appearance.BackColor = C
            '    .Columns("GRP_" & T).Header.Appearance.BackColor = C
            '    .Columns("SLS_PCT_" & T).Header.Appearance.BackColor = C
            '    .Columns("GRP_PCT_" & T).Header.Appearance.BackColor = C

            '    .Groups(T).Header.Appearance.BackColor = C

            '    .Columns("SLS_" & T).Width = W
            '    .Columns("CGS_" & T).Width = W
            '    .Columns("GRP_" & T).Width = W
            '    .Columns("SLS_PCT_" & T).Width = 60
            '    .Columns("GRP_PCT_" & T).Width = 60

            '    .Columns("SLS_" & T).Header.Caption = "Sales"
            '    .Columns("CGS_" & T).Header.Caption = "CGS"
            '    .Columns("GRP_" & T).Header.Caption = "$GP"
            '    .Columns("SLS_PCT_" & T).Header.Caption = "%Sls"
            '    .Columns("GRP_PCT_" & T).Header.Caption = "GP%"

            '    .Columns("SLS_" & T).Group = .Groups(T)
            '    .Columns("CGS_" & T).Group = .Groups(T)
            '    .Columns("GRP_" & T).Group = .Groups(T)
            '    .Columns("SLS_PCT_" & T).Group = .Groups(T)
            '    .Columns("GRP_PCT_" & T).Group = .Groups(T)

            'Next
        ' End With

        Call Create_Summary(grdSATSLSC1, "CUST_CODE", "Count")

        For i As Integer = 1 To 12
            Call Create_Summary(grdSATSLSC1, "GROSS_" & Format(i, "00"))
            Call Create_Summary(grdSATSLSC1, "RET_" & Format(i, "00"))
            Call Create_Summary(grdSATSLSC1, "NET_" & Format(i, "00"))
            Call Create_Summary(grdSATSLSC1, "RETL_" & Format(i, "00"))
        Next
        Call Create_Summary(grdSATSLSC1, "GROSS_TOT")
        Call Create_Summary(grdSATSLSC1, "RET_TOT")
        Call Create_Summary(grdSATSLSC1, "NET_TOT")
        Call Create_Summary(grdSATSLSC1, "RETL_TOT")

        grdSATSLSC1.DisplayLayout.UseFixedHeaders = True
        With grdSATSLSC1.DisplayLayout.Bands("SATSLSC1")
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("CUST_CLASS_CODE").Header.Fixed = True
            .Columns("MARGIN").Header.Fixed = True
        End With

        'Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP.ToString.Substring(0, 4) & "12"
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

        tabMain.Visible = tf
        'Setup_Summary()

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("SATSLSC1").Rows.Clear()

        dst.EnforceConstraints = True

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Reading from Sales History Data")
        Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)

        Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
        'z = Mid(z, 1, 4) & Mid(z, 6, 2)
        RYP = z
        'FYP = Mid(z, 1, 4) & "01"
        FYP = ASCMAIN1.Period_Calc(z, -11)

        Create_SATSLSC0(FYP, RYP)
        Call ASCMAIN1.Progress("Now Loading Data")

        dst.EnforceConstraints = False

        Fill_Records("SATSLSC1")
        Dim Sql As String = "Select CUST_CODE, CUST_NAME, CUST_CLASS_CODE, PRICE_CLASS_CODE, MARGIN" & vbCr

        For i As Integer = 0 To 11
            Dim p As String = ASCMAIN1.Period_Calc(FYP, +(i))
            'Sql &= ",sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' AND INV_TYPE = 'I' THEN NVL(INV_SALES,0) ELSE 0 END) GROSS_" & Format(i + 1, "00") & vbCr _
            '    & ",sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' AND INV_TYPE = 'C' THEN NVL(INV_SALES,0) ELSE 0 END) RET_" & Format(i + 1, "00") & vbCr _
            '    & ",sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' THEN NVL(INV_SALES,0) ELSE 0 END) NET_" & Format(i + 1, "00") & vbCr _
            '    & ",sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' THEN NVL(RETL_AMT,0) ELSE 0 END) RETL_" & Format(i + 1, "00") & vbCr
            Sql &= ",round(sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' AND INV_TYPE = 'I' THEN NVL(INV_SALES,0) ELSE 0 END)) GROSS_" & Format(i + 1, "00") & vbCr _
              & ",round(sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' AND INV_TYPE = 'C' THEN NVL(INV_SALES,0) ELSE 0 END)) RET_" & Format(i + 1, "00") & vbCr _
              & ",round(sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' THEN NVL(INV_SALES,0) ELSE 0 END)) NET_" & Format(i + 1, "00") & vbCr _
              & ",round(sum(CASE WHEN ORDR_YYYYPP_UPDATED = '" & p & "' THEN NVL(RETL_AMT,0) ELSE 0 END)) RETL_" & Format(i + 1, "00") & vbCr
        Next
        'Sql &= ",sum(CASE WHEN INV_TYPE = 'I' THEN NVL(INV_SALES,0) ELSE 0 END) GROSS_TOT" & vbCr _
        '    & ",sum(CASE WHEN INV_TYPE = 'C' THEN NVL(INV_SALES,0) ELSE 0 END) RET_TOT" & vbCr _
        '    & ",sum(NVL(INV_SALES,0)) NET_TOT" & vbCr _
        '    & ",sum(NVL(RETL_AMT,0)) RETL_TOT" & vbCr _
        Sql &= ",round(sum(CASE WHEN INV_TYPE = 'I' THEN NVL(INV_SALES,0) ELSE 0 END)) GROSS_TOT" & vbCr _
            & ",round(sum(CASE WHEN INV_TYPE = 'C' THEN NVL(INV_SALES,0) ELSE 0 END)) RET_TOT" & vbCr _
            & ",round(sum(NVL(INV_SALES,0))) NET_TOT" & vbCr _
            & ",round(sum(NVL(RETL_AMT,0))) RETL_TOT" & vbCr _
            & " from " & SATSLSC0 & " SATSLSC1 group by CUST_CODE, CUST_NAME, CUST_CLASS_CODE, PRICE_CLASS_CODE, MARGIN"
        Call Fill_Records("SATSLSC1", "", True, Sql)
        For i As Integer = 1 To 12
            Dim p As String = ASCMAIN1.Period_Calc(FYP, (i - 1))
            Dim caption As String = ASCMAIN1.Get_Legend(p).ToString.Substring(9, 6) & " $"
            With grdSATSLSC1.DisplayLayout.Bands(0).Groups("M_" & Format(i, "00"))
                .Header.Caption = caption
                .Header.Appearance.TextHAlign = HAlign.Center
                If i Mod 2 = 1 Then
                    .Header.Appearance.BackColor = Drawing.Color.Yellow
                Else
                    .Header.Appearance.BackColor = Drawing.Color.Cyan
                End If
            End With
        Next

        'EnforceConstraints(True)

        Call ASCMAIN1.Progress("Now Setting Up Screen")
        Setup_tabMain()

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
        Call Load_Popup_Menu(grdSATSLSC1, "SSB", "Show Filter", "Show GroupBox", "Customer Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case "grdSATSLSC1"
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
#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Create_SATSLSC0(ByVal FYP As String, ByVal RYP As String)

        ASCMAIN1.sql = "Select SOTINVH1.*, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CLASS_CODE" _
        & ", NVL(ARTCUST1.PRICE_CLASS_CODE,0) PRICE_CLASS_CODE, 100-NVL(ARTCUST1.PRICE_CLASS_CODE,0) MARGIN , 0 RETL_AMT" _
                 & " from SOTINVH1, ARTCUST1" & vbCr _
                 & " where SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE (+) " & vbCr _
                 & " and ORDR_YYYYPP_UPDATED between '" & FYP & "' and '" & RYP & "'" & vbCr _
                 & " and nvl(INV_SALES,0) <> 0"

        If SATSLSC0 = "" Then
            SATSLSC0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSC0 & " modify RETL_AMT NUMBER (12,2)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSC0)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSC0 & " " & ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Update " & SATSLSC0 & " set RETL_AMT=INV_SALES/" _
                       & "DECODE(100-NVL(PRICE_CLASS_CODE,0),0,1,((100-NVL(PRICE_CLASS_CODE,0))*.01)) ")
        End If


    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Call Print_Report_Begin()

        'SUBT = "Period Ending " & ASCMAIN1.Get_Legend(RYP).substring(9, 6) _
        'Generate_Report("SARSLSJ1", "Sales Analysis by Customer", SUBT)
        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ2", "Sales Analysis by Sales Rep", SUBT)
        'CR_params.Add("DETAIL", "N")
        'Generate_Report("SARSLSJ2", "Sales Analysis by Sales Rep-Summary", SUBT)
        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ3", "Sales Analysis by State", SUBT)
        'CR_params.Add("DETAIL", "N")
        'Generate_Report("SARSLSJ3", "Sales Analysis by State-Summary", SUBT)
        'Generate_Report("SARSLSJ4", "Sales Analysis by Customer-Rank", SUBT)
        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Call Print_Report_End()
    End Sub

    Private Sub optCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Setup_Summary()
    End Sub

    Sub Setup_Summary()
        If SELECTION_NO = 0 Then Exit Sub
        'tabSummary.Tabs("C").Visible = (optCode.Value = "C")
        'tabSummary.Tabs("R").Visible = (optCode.Value = "R")
        'tabSummary.Tabs("S").Visible = (optCode.Value = "S")
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        'If tabMain.SelectedTab Is Nothing Then
        '    UltraExplorerBar1.Groups("Summaries").Visible = False
        'Else
        '    UltraExplorerBar1.Groups("Summaries").Visible = (tabMain.SelectedTab.Key = "Summaries")
        '    Setup_Summary()
        'End If
    End Sub
    Private Sub grdSATSLSC1_AfterRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLSC1.InitializeRow
        Call Set_BackColor(sender, e)
    End Sub
    Sub Set_BackColor(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        If e.Row.IsDataRow Then
            Try
                For I As Int16 = 1 To 12
                    'If I Mod 2 = 1 Then
                    '    e.Row.Cells("GROSS_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Yellow
                    '    e.Row.Cells("RET_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Yellow
                    '    e.Row.Cells("NET_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Yellow
                    '    e.Row.Cells("RETL_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Yellow
                    'Else
                    '    e.Row.Cells("GROSS_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Cyan
                    '    e.Row.Cells("RET_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Cyan
                    '    e.Row.Cells("NET_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Cyan
                    '    e.Row.Cells("RETL_" & Format(I, "00")).Appearance.BackColor = Drawing.Color.Cyan
                    'End If
                Next
                'e.Row.Cells("GROSS_TOT").Appearance.BackColor = Drawing.Color.Cyan
                'e.Row.Cells("RET_TOT").Appearance.BackColor = Drawing.Color.Cyan
                'e.Row.Cells("NET_TOT").Appearance.BackColor = Drawing.Color.Cyan
                'e.Row.Cells("RETL_TOT").Appearance.BackColor = Drawing.Color.Cyan

            Catch ex As Exception
                ' Nothing
            End Try
        End If
    End Sub

End Class