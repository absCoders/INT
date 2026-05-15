Public Class SAFBUDW1

    Dim rowARTCUST1 As DataRow
    Dim BRAND_CODE As String
    Dim OPS_YYYY As String
    Dim MMM_CUR As String
    Dim MMM_NXT As String
    Dim in_000s As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ICTBRAN1", "*")

            ASCMAIN1.sql = "Select SATBUDW1.*" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", DECODE(INSTR(SATBUDW1.CUST_CODE,':'),0,ARTCUST1.TRADE_CLASS_CODE,SUBSTR(SATBUDW1.CUST_CODE,1,INSTR(SATBUDW1.CUST_CODE,':')-1)) TRADE_CLASS_CODE" & vbCrLf _
                & ", DECODE(INSTR(SATBUDW1.CUST_CODE,':'),0,ARTCUST1.CUST_CLASS_CODE,SUBSTR(SATBUDW1.CUST_CODE,INSTR(SATBUDW1.CUST_CODE,':')+1)) CUST_CLASS_CODE" & vbCrLf _
                & " from SATBUDW1,ARTCUST1 " & vbCrLf _
                & " where SATBUDW1.BRAND_CODE = :PARM1 " & vbCrLf _
                & "   and SATBUDW1.OPS_YYYY = :PARM2" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = SATBUDW1.CUST_CODE"
            Create_TDA(.Tables.Add, "SATBUDW1", "**", 0, True, "VV", 3)
            For I As Integer = 1 To 12
                For Each PFX As String In New String() {"OB", "RB", "WB"}
                    dst.Tables("SATBUDW1").Columns("OB_P" & Format(I, "00")).DataType = GetType(System.Decimal)
                Next
            Next

            Dim TOTAL As String = ""

            With .Tables("SATBUDW1").Columns
                For I As Integer = 1 To 12 + 4 + 2
                    For Each PFX As String In New String() {"TY", "OB", "RB", "WB", "LY", "TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP", "RT"}
                        If I > 12 Or (PFX <> "OB" And PFX <> "RB" And PFX <> "WB") Then
                            .Add(PFX & "_P" & Format(I, "00"), GetType(System.Decimal))
                        End If
                    Next

                    .Item("TYRBVD_P" & Format(I, "00")).Expression = Replace("ISNULL(TY_P00,0) - ISNULL(RB_P00,0)", "P00", "P" & Format(I, "00"))
                    .Item("TYRBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(RB_P00,0)=0,0,100*ISNULL(TYRBVD_P00,0)/ISNULL(RB_P00,0))", "P00", "P" & Format(I, "00"))
                    .Item("TYWBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(WB_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(WB_P00,0))/ISNULL(WB_P00,0))", "P00", "P" & Format(I, "00"))
                    .Item("TYLYVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(LY_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(LY_P00,0))/ISNULL(LY_P00,0))", "P00", "P" & Format(I, "00"))

                    If I <= 12 Then
                        TOTAL &= "+ISNULL(TY_P" & Format(I, "00") & ",0)"
                    End If
                Next
                '
                For Each PFX As String In New String() {"TY", "OB", "RB", "WB", "LY", "TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP", "RT"}

                    If PFX = "TYRBVP" Or PFX = "TYWBVP" Or PFX = "TYLYVP" Then
                        ' DO NOTHING - EXPRESSION WAS SET ABOVE
                        .Add(PFX & "_TOTAL", GetType(System.Decimal))
                        .Item(PFX & "_TOTAL").Expression = Replace(.Item(PFX & "_P01").Expression, "_P01", "_TOTAL")
                    Else
                        .Add(PFX & "_TOTAL", GetType(System.Decimal), Mid(Replace(TOTAL, "TY", PFX), 2))
                        .Item(PFX & "_P13").Expression = Mid(Replace("+ISNULL(TY_P01,0)+ISNULL(TY_P02,0)+ISNULL(TY_P03,0)", "TY", PFX), 2)
                        .Item(PFX & "_P14").Expression = Mid(Replace("+ISNULL(TY_P04,0)+ISNULL(TY_P05,0)+ISNULL(TY_P06,0)", "TY", PFX), 2)
                        .Item(PFX & "_P15").Expression = Mid(Replace("+ISNULL(TY_P07,0)+ISNULL(TY_P08,0)+ISNULL(TY_P09,0)", "TY", PFX), 2)
                        .Item(PFX & "_P16").Expression = Mid(Replace("+ISNULL(TY_P10,0)+ISNULL(TY_P11,0)+ISNULL(TY_P12,0)", "TY", PFX), 2)
                        .Item(PFX & "_P17").Expression = Mid(Replace("+ISNULL(TY_P13,0)+ISNULL(TY_P14,0)", "TY", PFX), 2)
                        .Item(PFX & "_P18").Expression = Mid(Replace("+ISNULL(TY_P15,0)+ISNULL(TY_P16,0)", "TY", PFX), 2)
                    End If
                Next

                For I As Integer = 1 To 9
                    .Add("BKG" & CStr(I), GetType(System.Decimal))
                Next
            End With
          
            ASCMAIN1.sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC, BUDGET_BY_CUST from SOTTCLS1"
            Create_TDA(.Tables.Add, "SOTTCLS1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from ARTCLAS1"
            Create_TDA(.Tables.Add, "ARTCLAS1", "**", 0, False, "", 1)
        End With

        grdSOTTCLS1.DataSource = dst.Tables("SOTTCLS1")
        grdSATBUDW1.DataSource = dst.Tables("SATBUDW1")

        Format_grdSATBUDW1()

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs

        Dim LEGEND As String
        LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.CYP)
        MMM_CUR = Mid(LEGEND, 10, 3)

        LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1))
        MMM_NXT = Mid(LEGEND, 10, 3)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("BRAND_CODE")

                If cmbOPS_YYYY.Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Year"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SATSLSW1", Absx1.txtFor("BRAND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                'Fill_Records("ARTCLAS1")
                'For Each row As DataRow In dst.Tables("SOTTCLS1").Select("")
                '    Dim BUDGET_BY_CUST As String = row.Item("BUDGET_BY_CUST") & ""
                '    If BUDGET_BY_CUST <> "1" Then
                '        Dim CUST_CLASS_CODE As String = row.Item("CUST_CLASS_CODE")
                '        Dim rowARTCLAS1 As DataRow = dst.Tables("ARTCLAS1").Rows.Find(CUST_CLASS_CODE)
                '        Dim CUST_CODE As String = rowARTCLAS1.Item("CUST_CODE") & ""
                '        If CUST_CODE = "" Then
                '            EMsg &= vbCr & "Customer Class " & CUST_CLASS_CODE & " does Not have a Budgetary Customer Defined"
                '        Else
                '            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                '            If rowARTCUST1 Is Nothing Then
                '                EMsg &= vbCr & "Invalid value defined for Budgetary Customer for Customer Class " & CUST_CLASS_CODE & " (" & CUST_CODE & ")"
                '            Else
                '                If rowARTCUST1.Item("CUST_CLASS_CODE") & "" <> CUST_CLASS_CODE Then
                '                    EMsg &= vbCr & "Budgetary Customer " & CUST_CODE & " defined for Customer Class " & CUST_CLASS_CODE & " is coded to Class (" & rowARTCUST1.Item("CUST_CLASS_CODE") & ")"
                '                End If
                '            End If
                '        End If
                '    End If
                'Next


            Case "Copy to Original"
                If MsgBox("Are you sure that you want to Copy the Working Budget to the Original Budget?", _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Case "Copy to Revised"
                If MsgBox("Are you sure that you want to Copy the Working Budget to the Revised Budget?", _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Copy to Original"
                Copy_Working("OB")
                MsgBox("Working Budget has been Copied to the Original Budget," & vbCrLf & " and to the Revised Budget, too", MsgBoxStyle.OkOnly, "Success")

            Case "Copy to Revised"
                Copy_Working("RB")
                MsgBox("Working Budget has been Copied to the Revised Budget", MsgBoxStyle.OkOnly, "Success")

            Case "Excel"

                If Not ScreenMode Then
                    Combined_Export()
                    Exit Sub
                Else
                    Export_Budget_to_Excel()
                End If

            Case "Report"
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Copy to Original").Visible = False
                    .Items("Copy to Revised").Visible = False
                    .Items("Report").Settings.Enabled = iScreenMode
                    .Items("Report").Visible = False
                    '.Items("Excel").Settings.Enabled = iScreenMode
                End With
                .Groups("Customers").Visible = False
                .Groups("Customer Classes").Visible = False
                .Groups("Trade Classes").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        chk000s.Visible = ScreenMode
        If ScreenMode Then Set_Read_Only_for_ctl(chk000s, False)

        grdSATBUDW1.Visible = ScreenMode
        If ScreenMode Then

            Fill_Records("ARTCLAS1")

            Fill_Records("SOTTCLS1")
            Sort_grdColumns(grdSOTTCLS1, "TRADE_CLASS_CODE")
            Setup_SATBUDW1_TRADE_CLASS_CODE()

            Show_Copy_Options()
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("SATBUDW1").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("BRAND_CODE").Text = ""
        If Absx1.cbeFor("OPS_YYYY").Value & "" = "" Then
            Absx1.cbeFor("OPS_YYYY").Value = Now.Year
        End If
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)
        BRAND_CODE = Absx1.txtFor("BRAND_CODE").Text
        OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Text

        EnforceConstraints(False)
        Fill_Records("SATBUDW1", New String() {HFs("BRAND_CODE"), HFs("OPS_YYYY")})

        'For Each row As DataRow In dst.Tables("SOTTCLS1").Select("")
        '    Dim BUDGET_BY_CUST As String = row.Item("BUDGET_BY_CUST") & ""
        '    Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE") & ""
        '    If BUDGET_BY_CUST = "1" Then
        '        Dim sqlw As String = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"
        '        For Each rowSATBUDW1 As DataRow In dst.Tables("SATBUDW1").Select(sqlw)
        '            rowSATBUDW1.Item("TRADE_CLASS_CODE") = rowSATBUDW1.Item("TRADE_CLASS_CODE_ARTCUST1")
        '            rowSATBUDW1.Item("CUST_CLASS_CODE") = rowSATBUDW1.Item("CUST_CLASS_CODE_ARTCUST1")
        '        Next
        '    End If
        'Next

        EnforceConstraints(True)

        For M As Integer = 1 To 12
            Dim D As Date = CDate(Format(M, "00") & "/01/" & OPS_YYYY)
            grdSATBUDW1.DisplayLayout.Bands(0).Groups(Format(M, "00")).Header.Caption = Format(D, "MMM/yy")

            With dst.Tables("SATBUDW1").Columns("TY_P" & Format(M, "00"))
                .Expression = ""
                .ReadOnly = False
            End With

            With grdSATBUDW1.DisplayLayout.Bands(0).Columns("WB_P" & Format(M, "00"))
                If OPS_YYYY & Format(M, "00") >= ASCMAIN1.CYP Then ' ASCMAIN1.CYM
                    .CellAppearance.BackColor = Drawing.Color.Yellow
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .CellAppearance.BackColor = Drawing.Color.Orange
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    ' temp for mg
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With

            grdSATBUDW1.DisplayLayout.Bands(0).Columns("OB_P" & Format(M, "00")).CellActivation = UltraWinGrid.Activation.NoEdit
            grdSATBUDW1.DisplayLayout.Bands(0).Columns("RB_P" & Format(M, "00")).CellActivation = UltraWinGrid.Activation.NoEdit

            '' temp for mg
            'grdSATBUDW1.DisplayLayout.Bands(0).Columns("OB_P" & Format(M, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
            'grdSATBUDW1.DisplayLayout.Bands(0).Columns("RB_P" & Format(M, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit

        Next

        Get_History()

        With dst.Tables("SATBUDW1")
            .Columns("BKG2").Expression = ""
            .Columns("BKG5").Expression = ""
            If OPS_YYYY = Mid(ASCMAIN1.CYP, 1, 4) Then
                .Columns("BKG2").Expression = "ISNULL(TY_P" & Mid(ASCMAIN1.CYP, 5, 2) & ",0) - ISNULL(BKG3,0) - ISNULL(BKG4,0)"
                ' need to subtract curr open and pick because they are added into the TY Actual column for the current month
                .Columns("BKG5").Expression = "ISNULL(BKG2,0)+ISNULL(BKG3,0)+ISNULL(BKG4,0)"
            End If
        End With

        If OPS_YYYY = Mid(ASCMAIN1.CYP, 1, 4) Then
            If Mid(ASCMAIN1.CYP, 5, 2) <> "12" Then
                For M As Integer = Val(Mid(ASCMAIN1.CYP, 5, 2)) + 1 To 12
                    With dst.Tables("SATBUDW1").Columns("TY_P" & Format(M, "00"))
                        .Expression = "WB_P" & Format(M, "00")
                    End With
                Next
            End If
        End If

        grdSATBUDW1.DisplayLayout.Bands(0).MaxRows = 1

        If chk000s.Checked Then
            Setup_SATBUDW1_000s(True)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'For Each row As DataRow In dst.Tables("SOTTCLS1").Select("")
        '    Dim BUDGET_BY_CUST As String = row.Item("BUDGET_BY_CUST") & ""
        '    Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE") & ""
        '    If BUDGET_BY_CUST <> "1" Then
        '        Dim sqlw As String = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"
        '        For Each rowSATBUDW1 As DataRow In dst.Tables("SATBUDW1").Select(sqlw)
        '            rowSATBUDW1.Item("CUST_CODE") = "*"
        '        Next
        '    End If
        'Next

        Dim sql_Delete As String = "Delete from SATBUDW1" _
            & " where OPS_YYYY = '" & HFs("OPS_YYYY") & "'" _
            & " and BRAND_CODE = '" & HFs("BRAND_CODE") & "'"
        Update_Record_TDA("SATBUDW1", sql_Delete)

        If chk000s.Checked Then
            Setup_SATBUDW1_000s(False)
        End If

        CommitTrans("Update Complete")

    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSATBUDW1, "SSS", "Show Filter", "Show GroupBox", "Hide Qtr/Half")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdDPTFCSTD"
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Hide Qtr/Half"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                For M As Integer = 13 To 18
                    grd.DisplayLayout.Bands(0).Groups(Format(M, "00")).Hidden = tlb_sbt.Checked
                Next
        End Select

        Select Case grd.Name
            'Case "grdDPTFCSTD"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Call Click_Command("Load", e)
                End If

            Case "CUST_CLASS_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim CUST_CLASS_CODE As String = Absx1.txtFor("CUST_CLASS_CODE").Text
                    Add_CUST_CLASS_CODE(CUST_CLASS_CODE)
                End If

            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Add_CUST_CODE(CUST_CODE)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BRAND_CODE"
                'Call Click_Command("Load")

            Case "CUST_CODE"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Add_CUST_CODE(CUST_CODE)

            Case "CUST_CLASS_CODE"
                Dim CUST_CLASS_CODE As String = Absx1.txtFor("CUST_CLASS_CODE").Text
                Add_CUST_CLASS_CODE(CUST_CLASS_CODE)
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdSATBUDW1"

    Private Sub grdSATBUDW1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSATBUDW1.AfterCellUpdate

    End Sub

    Private Sub grdSATBUDW1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSATBUDW1.AfterRowActivate

    End Sub

    Private Sub grdSATBUDW1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSATBUDW1.AfterRowUpdate
    End Sub

    Private Sub grdSATBUDW1_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdSATBUDW1.BeforeCellActivate
    End Sub

    Private Sub grdSATBUDW1_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSATBUDW1.BeforeCellUpdate
    End Sub

    Private Sub grdSATBUDW1_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSATBUDW1.BeforeExitEditMode

    End Sub

    Private Sub grdSATBUDW1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSATBUDW1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Val(grow.Cells("TY_TOTAL").Value & "") <> 0 Or Val(grow.Cells("LY_TOTAL").Value & "") <> 0 Or Val(grow.Cells("RT_TOTAL").Value & "") <> 0 _
            Or Val(grow.Cells("BKG2").Value & "") <> 0 Or Val(grow.Cells("BKG3").Value & "") <> 0 _
            Or Val(grow.Cells("BKG7").Value & "") <> 0 Or Val(grow.Cells("BKG8").Value & "") <> 0 Then
                MsgBox("Cannot Delete " & grow.Cells("CUST_CODE").Value & "; History or Open Orders", MsgBoxStyle.OkOnly, "De-Select this Row - Delete not Permitted")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Private Sub grdSATBUDW1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSATBUDW1.BeforeRowUpdate

    End Sub

    Private Sub grdSATBUDW1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSATBUDW1.ClickCellButton
        'Dim sql_where As String = ""
        'Call grdClickCellButton(grdSATBUDW1, sql_where, False)
    End Sub

    Private Sub grdSATBUDW1_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSATBUDW1.Error
        grdSATBUDW1.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdSATBUDW1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATBUDW1.InitializeRow
        e.Row.Cells("DATA_01").Value = "TY-Act+Work"
        e.Row.Cells("DATA_02").Value = "Original Budget"
        e.Row.Cells("DATA_03").Value = "Revised Budget"
        e.Row.Cells("DATA_04").Value = "Working Budget"
        e.Row.Cells("DATA_05").Value = "LY " & Format(Val(OPS_YYYY) - 1, "0000")
        e.Row.Cells("DATA_06").Value = "$Var TY/Revd"
        e.Row.Cells("DATA_07").Value = "%Var TY/Revd"
        e.Row.Cells("DATA_08").Value = "% TY-Act/Work"
        e.Row.Cells("DATA_09").Value = "% TY-Act/LY"
        e.Row.Cells("DATA_10").Value = "Returns " & Format(Val(OPS_YYYY) - 1, "0000")

        e.Row.Cells("BKG2_DESC").Value = "Shipped MTD " & MMM_CUR
        e.Row.Cells("BKG3_DESC").Value = MMM_CUR & " Released"
        e.Row.Cells("BKG4_DESC").Value = MMM_CUR & " Open"
        e.Row.Cells("BKG5_DESC").Value = "Total Projected"

        e.Row.Cells("BKG7_DESC").Value = MMM_NXT & " Released"
        e.Row.Cells("BKG8_DESC").Value = MMM_NXT & " Open"

    End Sub
#End Region

    Sub Format_grdSATBUDW1()

        With grdSATBUDW1.DisplayLayout.Bands(0)


            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key.StartsWith("WB_P") Then
                    gcol.CellAppearance.BackColor = System.Drawing.Color.Yellow
                Else
                    gcol.CellAppearance.BackColor = System.Drawing.Color.Beige
                End If
            Next

            Dim lvl As Integer

            .LevelCount = 10
            .ColHeadersVisible = False

            Dim G As UltraWinGrid.UltraGridGroup

            G = .Groups.Add("CUST_CODE")
            G.Header.Caption = "Customer"
            G.Header.Fixed = True
            G.Header.Appearance.BackColor2 = Drawing.Color.Orange
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            With .Columns("CUST_CODE")
                .Group = G
                .Width = 100
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .CellAppearance.BackColor = Drawing.Color.Beige
                Create_Summary(grdSATBUDW1, "CUST_CODE", "Count")
            End With

            lvl = 0
            For b As Integer = 1 To 9
                Dim C As String = "BKG" & CStr(b)
                With .Columns(C)
                    .Group = G
                    lvl += 1
                    .Level = lvl
                    .Format = "#,##0"
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .CellAppearance.BackColor = Drawing.Color.Ivory
                    Create_Summary(grdSATBUDW1, C)
                End With
            Next

            G = .Groups.Add("CUST_NAME")
            G.Header.Caption = "Customer Name"
            G.Header.Fixed = True
            G.Header.Appearance.BackColor2 = Drawing.Color.Orange
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            With .Columns("CUST_NAME")
                .Group = G
                .Width = 200
                .CellActivation = UltraWinGrid.Activation.NoEdit
                Create_Summary(grdSATBUDW1, "CUST_NAME", "Count")
            End With

            lvl = 0
            For b As Integer = 1 To 9
                Dim C As String = "BKG" & CStr(b) & "_DESC"
                With .Columns.Add(C)
                    .Group = G
                    lvl += 1
                    .Level = lvl
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .CellAppearance.BackColor = Drawing.Color.Ivory
                    Create_Summary(grdSATBUDW1, C, "Min")
                End With
            Next

            G = .Groups.Add("TRADE_CLASS_CODE")
            G.Header.Caption = "TrCls"
            G.Header.Fixed = True
            G.Header.Appearance.BackColor2 = Drawing.Color.Orange
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            With .Columns("TRADE_CLASS_CODE")
                .Group = G
                .Width = 50
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End With
            G.Hidden = True

            G = .Groups.Add("DATA")
            G.Header.Fixed = True
            G.Header.Caption = "Data"
            G.Header.Appearance.BackColor2 = Drawing.Color.Yellow
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            For I As Integer = 1 To 10
                With .Columns.Add("DATA_" & Format(I, "00"))
                    .Group = G
                    .Level = I - 1
                    .Width = 150
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .CellAppearance.BackColor = Drawing.Color.Beige
                    Create_Summary(grdSATBUDW1, "DATA_" & Format(I, "00"), "Min")
                    If I >= 6 And I <= 9 Then
                        .CellAppearance.BackColor = Drawing.Color.LightGray
                    ElseIf I = 1 Then
                        .CellAppearance.BackColor = Drawing.Color.LightGreen
                    ElseIf I = 5 Then
                        .CellAppearance.BackColor = Drawing.Color.LightBlue
                    ElseIf I = 10 Then
                        .CellAppearance.BackColor = Drawing.Color.LightPink
                    End If
                End With
            Next



            lvl = 0
            G = .Groups.Add("TOTALS")
            G.Header.Caption = "Totals"
            G.Header.Appearance.TextHAlign = HAlign.Right
            G.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            G.Header.Fixed = True
            For Each P As String In New String() {"TY", "OB", "RB", "WB", "LY", "TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP", "RT"}
                Dim C As String = P & "_TOTAL"
                With .Columns(C)
                    .Group = G
                    .Width = 80
                    .Level = lvl
                    lvl += 1
                    .Format = "#,##0"
                    If P = "TYRBVP" Or P = "TYWBVP" Or P = "TYLYVP" Then
                        .Format = "#,##0.00"
                    End If

                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    If P = "TYRBVP" Or P = "TYWBVP" Or P = "TYLYVP" Then
                        Create_Summary(grdSATBUDW1, C, "Custom")
                    Else
                        Create_Summary(grdSATBUDW1, C)
                    End If


                    If New String() {"TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP"}.Contains(P) Then
                        .CellAppearance.BackColor = Drawing.Color.LightGray
                    End If
                    If New String() {"RT"}.Contains(P) Then
                        .CellAppearance.BackColor = Drawing.Color.LightPink
                    End If
                End With
            Next

            For Each M As Integer In New Integer() {1, 2, 3, 13, 4, 5, 6, 14, 17, 7, 8, 9, 15, 10, 11, 12, 16, 18}
                lvl = 0
                G = .Groups.Add(Format(M, "00"))
                If M <= 12 Then
                    G.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf M <= 16 Then
                    G.Header.Appearance.BackColor2 = Drawing.Color.Khaki
                    G.Header.Caption = CStr(M - 12) & "Q"
                Else
                    G.Header.Appearance.BackColor2 = Drawing.Color.PeachPuff
                    G.Header.Caption = CStr(M - 16) & "H"
                End If
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                G.Header.Appearance.TextHAlign = HAlign.Right
                For Each P As String In New String() {"TY", "OB", "RB", "WB", "LY", "TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP", "RT"}
                    Dim C As String = P & "_P" & Format(M, "00")
                    With .Columns(C)
                        If M <= 12 Then
                        ElseIf M <= 16 Then
                            .CellAppearance.BackColor = Drawing.Color.Khaki
                        Else
                            .CellAppearance.BackColor = Drawing.Color.LightGoldenrodYellow
                        End If
                        .Group = G
                        .Width = 80
                        .Level = lvl
                        lvl += 1
                        .Format = "#,##0"
                        If P = "TYRBVP" Or P = "TYWBVP" Or P = "TYLYVP" Then
                            .Format = "#,##0.00"
                        End If

                        If P = "WB" And M < 12 Then
                            .CellActivation = UltraWinGrid.Activation.AllowEdit
                        Else
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                        End If

                        If P = "TYRBVP" Or P = "TYWBVP" Or P = "TYLYVP" Then
                            Create_Summary(grdSATBUDW1, C, "Custom")
                        Else
                            Create_Summary(grdSATBUDW1, C)
                        End If


                        If New String() {"TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP"}.Contains(P) Then
                            .CellAppearance.BackColor = Drawing.Color.LightGray
                        End If
                        If New String() {"RT"}.Contains(P) Then
                            .CellAppearance.BackColor = Drawing.Color.LightPink
                        End If
                    End With
                Next
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In grdSATBUDW1.DisplayLayout.Bands(0).Columns
                gcol.TabIndex = 0
                gcol.TabStop = False
            Next
            For M As Integer = 1 To 12
                grdSATBUDW1.DisplayLayout.Bands(0).Columns("WB_P" & Format(M, "00")).TabIndex = M - 1
                grdSATBUDW1.DisplayLayout.Bands(0).Columns("WB_P" & Format(M, "00")).TabStop = True
            Next
        End With
    End Sub

    Sub Get_History()

        Dim sql As String = ""
        For I As Integer = 1 To 12
            sql &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & Format(Val(OPS_YYYY) - 1, "0000") & Format(I, "00") & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) LY_P" & Format(I, "00") & vbCrLf
            sql &= ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & OPS_YYYY & Format(I, "00") & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_P" & Format(I, "00") & vbCrLf
        Next

        ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE" & vbCrLf _
            & ", NVL(ARTCUST1.TRADE_CLASS_CODE,'?') TRADE_CLASS_CODE" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CLASS_CODE,'?') CUST_CLASS_CODE" & vbCrLf _
            & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',SOTINVH2.CUST_CODE,NVL(ARTCUST1.TRADE_CLASS_CODE,'?') || ':' || NVL(ARTCUST1.CUST_CLASS_CODE,'?')) CUST_CODE" & vbCrLf _
            & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',ARTCUST1.CUST_NAME,ARTCLAS1.CUST_CLASS_DESC) CUST_NAME" & vbCrLf _
            & sql _
            & " from SOTINVH2,ARTCUST1,ICTITEM1,ICTCOLL1,SOTTCLS1,ARTCLAS1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTCOLL1.BRAND_CODE = '" & HFs("BRAND_CODE") & "'" & vbCrLf _
            & "   and ARTCLAS1.CUST_CLASS_CODE (+) = ARTCUST1.CUST_CLASS_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & Format(Val(OPS_YYYY) - 1, "0000") & "01'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & OPS_YYYY & "12'" & vbCrLf _
            & " group by SOTINVH2.INV_TYPE" & vbCrLf _
            & ", NVL(ARTCUST1.TRADE_CLASS_CODE,'?')" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CLASS_CODE,'?')" & vbCrLf _
            & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',SOTINVH2.CUST_CODE,NVL(ARTCUST1.TRADE_CLASS_CODE,'?') || ':' || NVL(ARTCUST1.CUST_CLASS_CODE,'?'))  " & vbCrLf _
            & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',ARTCUST1.CUST_NAME,ARTCLAS1.CUST_CLASS_DESC)  "

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim INV_TYPE As String = row.Item("INV_TYPE") & ""
            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE") & ""
            Dim CUST_CLASS_CODE As String = row.Item("CUST_CLASS_CODE") & ""
            Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
            Dim CUST_NAME As String = row.Item("CUST_NAME") & ""
            Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find(New String() {OPS_YYYY, BRAND_CODE, CUST_CODE})
            If rowSATBUDW1 Is Nothing Then
                rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                rowSATBUDW1.ITEM("OPS_YYYY") = OPS_YYYY
                rowSATBUDW1.ITEM("BRAND_CODE") = BRAND_CODE
                rowSATBUDW1.Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                rowSATBUDW1.Item("CUST_CLASS_CODE") = CUST_CLASS_CODE
                rowSATBUDW1.ITEM("CUST_CODE") = CUST_CODE
                rowSATBUDW1.ITEM("CUST_NAME") = CUST_NAME
                dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
            End If
            If INV_TYPE = "I" Then
                For i As Integer = 1 To 12
                    rowSATBUDW1.Item("TY_P" & Format(i, "00")) = row.Item("TY_P" & Format(i, "00"))
                    rowSATBUDW1.Item("LY_P" & Format(i, "00")) = row.Item("LY_P" & Format(i, "00"))
                Next
            Else
                For i As Integer = 1 To 12
                    rowSATBUDW1.Item("RT_P" & Format(i, "00")) = row.Item("LY_P" & Format(i, "00"))
                Next

            End If
        Next

        If OPS_YYYY = Mid(ASCMAIN1.CYP, 1, 4) Then
            sql = ""
            sql &= ", SUM (CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') <= '" & ASCMAIN1.CYP & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) CURR_OPEN" & vbCrLf
            sql &= ", SUM (CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') <= '" & ASCMAIN1.CYP & "' THEN NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) CURR_PICK" & vbCrLf
            sql &= ", SUM (CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') =  '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1) & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) NEXT_OPEN" & vbCrLf
            sql &= ", SUM (CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') =  '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1) & "' THEN NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0) ELSE 0 END) NEXT_PICK" & vbCrLf

            ASCMAIN1.sql = "Select 'O' INV_TYPE" & vbCrLf _
                & ", NVL(ARTCUST1.TRADE_CLASS_CODE,'?') TRADE_CLASS_CODE" & vbCrLf _
                & ", NVL(ARTCUST1.CUST_CLASS_CODE,'?') CUST_CLASS_CODE" & vbCrLf _
                & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',SOTORDR2.CUST_CODE,NVL(ARTCUST1.TRADE_CLASS_CODE,'?') || ':' || NVL(ARTCUST1.CUST_CLASS_CODE,'?')) CUST_CODE" & vbCrLf _
                & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',ARTCUST1.CUST_NAME,ARTCLAS1.CUST_CLASS_DESC) CUST_NAME" & vbCrLf _
                & sql _
                & " from SOTORDR2,SOTORDR1,ARTCUST1,ICTITEM1,ICTCOLL1,SOTTCLS1,ARTCLAS1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ARTCLAS1.CUST_CLASS_CODE (+) = ARTCUST1.CUST_CLASS_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ICTCOLL1.BRAND_CODE = '" & HFs("BRAND_CODE") & "'" & vbCrLf _
                & "   and SOTORDR2.ORDR_STATUS >= 'O'" & vbCrLf _
                & "   and SOTORDR2.ORDR_STATUS <= 'P'" & vbCrLf _
                & " group by 'O'" & vbCrLf _
                & ", NVL(ARTCUST1.TRADE_CLASS_CODE,'?')" & vbCrLf _
                & ", NVL(ARTCUST1.CUST_CLASS_CODE,'?')" & vbCrLf _
                & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',SOTORDR2.CUST_CODE,NVL(ARTCUST1.TRADE_CLASS_CODE,'?') || ':' || NVL(ARTCUST1.CUST_CLASS_CODE,'?'))  " & vbCrLf _
                & ", DECODE(NVL(SOTTCLS1.BUDGET_BY_CUST,'0'),'1',ARTCUST1.CUST_NAME,ARTCLAS1.CUST_CLASS_DESC)  "


            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE") & ""
                Dim CUST_CLASS_CODE As String = row.Item("CUST_CLASS_CODE") & ""
                Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                Dim CUST_NAME As String = row.Item("CUST_NAME") & ""
                Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find(New String() {OPS_YYYY, BRAND_CODE, CUST_CODE})
                If rowSATBUDW1 Is Nothing Then
                    rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                    rowSATBUDW1.Item("OPS_YYYY") = OPS_YYYY
                    rowSATBUDW1.Item("BRAND_CODE") = BRAND_CODE
                    rowSATBUDW1.Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                    rowSATBUDW1.Item("CUST_CLASS_CODE") = CUST_CLASS_CODE
                    rowSATBUDW1.Item("CUST_CODE") = CUST_CODE
                    rowSATBUDW1.Item("CUST_NAME") = CUST_NAME
                    dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
                End If
                rowSATBUDW1.Item("BKG3") = row.Item("CURR_PICK")
                rowSATBUDW1.Item("BKG4") = row.Item("CURR_OPEN")
                rowSATBUDW1.Item("BKG7") = row.Item("NEXT_PICK")
                rowSATBUDW1.Item("BKG8") = row.Item("NEXT_OPEN")

                rowSATBUDW1.Item("TY_P" & Mid(ASCMAIN1.CYP, 5, 2)) = Val(rowSATBUDW1.Item("TY_P" & Mid(ASCMAIN1.CYP, 5, 2)) & "") _
                    + Val(row.Item("CURR_PICK") & "") + Val(row.Item("CURR_OPEN") & "")

            Next
        End If
    End Sub

    Sub Setup_SATBUDW1_TRADE_CLASS_CODE()
        Dim TRADE_CLASS_CODE As String = grdSOTTCLS1.ActiveRow.Cells("TRADE_CLASS_CODE").Value
        Dim TRADE_CLASS_DESC As String = grdSOTTCLS1.ActiveRow.Cells("TRADE_CLASS_DESC").Value & ""
        Dim dvw As DataView = DirectCast(grdSATBUDW1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"
        Dim BUDGET_BY_CUST As String = grdSOTTCLS1.ActiveRow.Cells("BUDGET_BY_CUST").Value & ""

        If EntryMode = "X" Then
            grdSATBUDW1.Text = Absx1.txtFor("BRAND_CODE").Text & " " & TRADE_CLASS_CODE & " - " & TRADE_CLASS_DESC
        Else
            grdSATBUDW1.Text = "Sales Budgets by " & IIf(BUDGET_BY_CUST = "1", "Customer", "Customer Class") & ", Trade Class " & TRADE_CLASS_CODE & " - " & TRADE_CLASS_DESC
        End If

        Sort_grdColumns(grdSATBUDW1, "CUST_CODE")

        grdSATBUDW1.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for " & TRADE_CLASS_CODE & " - " & TRADE_CLASS_DESC

        UltraExplorerBar1.Groups("Customers").Visible = (BUDGET_BY_CUST = "1")
        UltraExplorerBar1.Groups("Customer Classes").Visible = Not (BUDGET_BY_CUST = "1")
    End Sub

    Private Sub grdSOTTCLS1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTTCLS1.AfterRowActivate
        If SELECTION_NO = 0 Then Exit Sub
        If EntryMode = "" Then Exit Sub
        Setup_SATBUDW1_TRADE_CLASS_CODE()
    End Sub

    Sub Add_CUST_CODE(CUST_CODE As String)
        If CUST_CODE <> "" Then
            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
            If rowARTCUST1 Is Nothing Then
                MsgBox("Invalid Value Specified for Customer Code (" & CUST_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer")
            Else
                Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                If TRADE_CLASS_CODE <> grdSOTTCLS1.ActiveRow.Cells("TRADE_CLASS_CODE").Value Then
                    MsgBox("Customer Code (" & CUST_CODE & ") does not belong to Active Trade Class (" & grdSOTTCLS1.ActiveRow.Cells("TRADE_CLASS_CODE").Value & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                    Exit Sub
                End If

                Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find(New String() {OPS_YYYY, BRAND_CODE, CUST_CODE})
                If rowSATBUDW1 IsNot Nothing Then
                    MsgBox("Customer " & CUST_CODE & " is already in Budget", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                Else
                    rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                    rowSATBUDW1.Item("OPS_YYYY") = OPS_YYYY
                    rowSATBUDW1.Item("BRAND_CODE") = BRAND_CODE
                    rowSATBUDW1.Item("CUST_CODE") = CUST_CODE
                    rowSATBUDW1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                    rowSATBUDW1.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
                    rowSATBUDW1.Item("CUST_CLASS_CODE") = rowARTCUST1.Item("CUST_CLASS_CODE")
                    dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
                End If
                Absx1.txtFor("CUST_CODE").Text = ""

                For Each grow As UltraWinGrid.UltraGridRow In grdSATBUDW1.Rows
                    If grow.Cells("CUST_CODE").Value = CUST_CODE Then
                        grdSATBUDW1.ActiveRow = grow
                    End If
                Next

            End If
            Application.DoEvents()
            Absx1.txtFor("CUST_CODE").Focus()
        End If
    End Sub

    Sub Add_CUST_CLASS_CODE(CUST_CLASS_CODE As String)
        If CUST_CLASS_CODE <> "" Then
            Dim rowARTCLAS1 As DataRow = LookUp("ARTCLAS1", CUST_CLASS_CODE)
            If rowARTCLAS1 Is Nothing Then
                MsgBox("Invalid Value Specified for Customer Class Code (" & CUST_CLASS_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer Class")
            Else
                Dim TRADE_CLASS_CODE As String = grdSOTTCLS1.ActiveRow.Cells("TRADE_CLASS_CODE").Value
                Dim CUST_CODE As String = TRADE_CLASS_CODE & ":" & CUST_CLASS_CODE  ' "*"
                Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find(New String() {OPS_YYYY, BRAND_CODE, CUST_CODE})
                If rowSATBUDW1 IsNot Nothing Then
                    MsgBox("Customer " & CUST_CODE & " is already in Budget", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                Else
                    rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                    rowSATBUDW1.Item("OPS_YYYY") = OPS_YYYY
                    rowSATBUDW1.Item("BRAND_CODE") = BRAND_CODE
                    rowSATBUDW1.Item("CUST_CODE") = CUST_CODE
                    rowSATBUDW1.Item("CUST_NAME") = TRADE_CLASS_CODE & ":" & CUST_CLASS_CODE & " " & rowARTCLAS1.Item("CUST_CLASS_DESC")
                    rowSATBUDW1.Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                    rowSATBUDW1.Item("CUST_CLASS_CODE") = CUST_CLASS_CODE
                    dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
                End If
                Absx1.txtFor("CUST_CLASS_CODE").Text = ""

                For Each grow As UltraWinGrid.UltraGridRow In grdSATBUDW1.Rows
                    If grow.Cells("CUST_CODE").Value = CUST_CODE Then
                        grdSATBUDW1.ActiveRow = grow
                    End If
                Next

            End If
            Application.DoEvents()
            Absx1.txtFor("CUST_CLASS_CODE").Focus()
        End If
    End Sub

    Sub Show_Copy_Options()
        Dim o As Int64 = Val(dst.Tables("SATBUDW1").Compute("SUM(OB_TOTAL)", "") & "")
        UltraExplorerBar1.Groups("Screen Control").Items("Copy to Original").Visible = (o = 0)
        UltraExplorerBar1.Groups("Screen Control").Items("Copy to Revised").Visible = Not (o = 0)
    End Sub

    Sub Copy_Working(Copy_to As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Copying Working Budget", "")

        For Each row As DataRow In dst.Tables("SATBUDW1").Select("")
            For I As Integer = 1 To 12
                row.Item(Copy_to & "_P" & Format(I, "00")) = row.Item("WB_P" & Format(I, "00"))
                If Copy_to = "OB" Then
                    row.Item("RB_P" & Format(I, "00")) = row.Item("WB_P" & Format(I, "00"))
                End If
            Next
        Next
        Show_Copy_Options()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub chk000s_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chk000s.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_SATBUDW1_000s(chk000s.Checked)
    End Sub

    Sub Setup_SATBUDW1_000s(divide_by_000 As Boolean)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Converting 000s")

        Dim FS As New Dictionary(Of String, String)
        For I As Integer = 1 To 18
            For Each PFX As String In New String() {"TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP"}
                Dim C As String = PFX & "_P" & Format(I, "00")
                FS.Add(C, dst.Tables("SATBUDW1").Columns(C).Expression)
                dst.Tables("SATBUDW1").Columns(C).Expression = ""
            Next
        Next


        Dim F As Decimal = IIf(divide_by_000, 0.001, 1000)
        Dim FMT As String = IIf(divide_by_000, "#,##0.0", "#,##0")
        Dim FMTD As String = "{0:" & FMT & "}"

        For Each row As DataRow In dst.Tables("SATBUDW1").Select("")
            For Each PFX As String In New String() {"TY", "OB", "RB", "WB", "LY", "TYRBVD", "RT"}
                For I As Integer = 1 To 18
                    Dim C As String = PFX & "_P" & Format(I, "00")
                    If Not dst.Tables("SATBUDW1").Columns(C).ReadOnly Then row.Item(C) = Val(row.Item(C) & "") * F
                    grdSATBUDW1.DisplayLayout.Bands(0).Columns(C).Format = FMT
                    grdSATBUDW1.DisplayLayout.Bands(0).Summaries(C).DisplayFormat = FMTD
                Next

                grdSATBUDW1.DisplayLayout.Bands(0).Columns(PFX & "_TOTAL").Format = FMT
                grdSATBUDW1.DisplayLayout.Bands(0).Summaries(PFX & "_TOTAL").DisplayFormat = FMTD
            Next

            For I As Integer = 1 To 9
                Dim C As String = "BKG" & Format(I, "0")
                If Not dst.Tables("SATBUDW1").Columns(C).ReadOnly Then row.Item(C) = Val(row.Item(C) & "") * F
                grdSATBUDW1.DisplayLayout.Bands(0).Columns(C).Format = FMT
                grdSATBUDW1.DisplayLayout.Bands(0).Summaries(C).DisplayFormat = FMTD
            Next
        Next

        For I As Integer = 1 To 18
            For Each PFX As String In New String() {"TYRBVD", "TYRBVP", "TYWBVP", "TYLYVP"}
                Dim C As String = PFX & "_P" & Format(I, "00")
                dst.Tables("SATBUDW1").Columns(C).Expression = FS(C)
            Next
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        '.Item("TYRBVD_P" & Format(I, "00")).Expression = Replace("ISNULL(TY_P00,0) - ISNULL(RB_P00,0)", "P00", "P" & Format(I, "00"))
        '.Item("TYRBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(RB_P00,0)=0,0,100*ISNULL(TYRBVD_P00,0)/ISNULL(RB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYWBVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(WB_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(WB_P00,0))/ISNULL(WB_P00,0))", "P00", "P" & Format(I, "00"))
        '.Item("TYLYVP_P" & Format(I, "00")).Expression = Replace("IIF(ISNULL(LY_P00,0)=0,0,100*(ISNULL(TY_P00,0)-ISNULL(LY_P00,0))/ISNULL(LY_P00,0))", "P00", "P" & Format(I, "00"))


        Select Case grd.Name
            Case "grdSATBUDW1"
                Dim KEY As String = summarySettings.Key
                If KEY.StartsWith("TYRBVP") Then
                    Dim RB As String = "RB" & Mid(KEY, 7)
                    Dim D As String = "TYRBVD" & Mid(KEY, 7)
                    TOTALS.Add(RB, 0)
                    TOTALS.Add(D, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(RB) <> 0 Then CustomValue = 100 * TOTALS(D) / TOTALS(RB)

                ElseIf KEY.StartsWith("TYWBVP") Then
                    Dim TY As String = "TY" & Mid(KEY, 7)
                    Dim WB As String = "WB" & Mid(KEY, 7)
                    TOTALS.Add(TY, 0)
                    TOTALS.Add(WB, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(WB) <> 0 Then CustomValue = 100 * (TOTALS(TY) - TOTALS(WB)) / TOTALS(WB)

                ElseIf KEY.StartsWith("TYLYVP") Then
                    Dim TY As String = "TY" & Mid(KEY, 7)
                    Dim LY As String = "LY" & Mid(KEY, 7)
                    TOTALS.Add(TY, 0)
                    TOTALS.Add(LY, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(LY) <> 0 Then CustomValue = 100 * (TOTALS(TY) - TOTALS(LY)) / TOTALS(LY)
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub CustomSummary_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                CustomSummary_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.StartsWith("TYRBVP") Then
                    Dim RB As String = "RB" & Mid(KEY, 7)
                    Dim D As String = "TYRBVD" & Mid(KEY, 7)
                    TOTALS(RB) += Val(grow2.Cells(RB).Value & "")
                    TOTALS(D) += Val(grow2.Cells(D).Value & "")

                ElseIf KEY.StartsWith("TYWBVP") Then
                    Dim TY As String = "TY" & Mid(KEY, 7)
                    Dim WB As String = "WB" & Mid(KEY, 7)
                    TOTALS(TY) += Val(grow2.Cells(TY).Value & "")
                    TOTALS(WB) += Val(grow2.Cells(WB).Value & "")

                ElseIf KEY.StartsWith("TYLYVP") Then
                    Dim TY As String = "TY" & Mid(KEY, 7)
                    Dim LY As String = "LY" & Mid(KEY, 7)
                    TOTALS(TY) += Val(grow2.Cells(TY).Value & "")
                    TOTALS(LY) += Val(grow2.Cells(LY).Value & "")
                End If
            End If
        Next
    End Sub

    Sub Export_Budget_to_Excel( _
        Optional multiple_brands As Boolean = False, _
        Optional ByRef r As Integer = 0, _
        Optional ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile = Nothing)

        Dim EntryMode_Save = EntryMode
        EntryMode = "X"

        Dim wsi As New List(Of Integer)

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text

        Me.Cursor = Cursors.WaitCursor

        'Dim myWorkbook As GemBox.Spreadsheet.ExcelFile = Nothing
        Dim myWorksheet As GemBox.Spreadsheet.ExcelWorksheet = Nothing

        'Dim r As Integer = 0
        grdSOTTCLS1.ActiveRow = Nothing
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTTCLS1.Rows
            grow.Activate()
            ASCMAIN1.Progress(grdSATBUDW1.Text)
            Application.DoEvents()
            If r = 0 Then
                'Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSATBUDW1, False, grdSATBUDW1.Text)
                myWorkbook = Gembox_Export_to_Excel(grdSATBUDW1, False, False, "")
                myWorksheet = myWorkbook.Worksheets(0) ' .Add("Sheet1")
            Else
                'myWorksheet = myWorkbook.Worksheets.Add("Sheet" & CStr(myWorkbook.Worksheets.Count + 1))
                'Gembox_Export_to_Excel_Add_grd_to_Sheet(grdSATBUDW1, r, 0, myWorksheet)
                Gembox_Export_to_Excel_Add_grd(myWorkbook, grdSATBUDW1, False, "", "A")

            End If

            wsi.Add(r)

            With myWorkbook.Worksheets(r).PrintOptions
                ' Necessary since upgrade to 3.7 since the default causes A4 to be used
                .PaperType = GemBox.Spreadsheet.PaperType.Letter
                .FitWorksheetWidthToPages = 1
                .Portrait = False
                .PrintGridlines = True

                .BottomMargin = 0.25
                .TopMargin = 0.25
                .LeftMargin = 0.25
                .RightMargin = 0.25
            End With

            r += 1
        Next
        ASCMAIN1.Progress("", "")

        Create_Budget_Summary(myWorkbook, BRAND_CODE, Absx1.txtFor("BRAND_NAME").Text, wsi)
        r += 1

        If Not multiple_brands Then Gembox_Export_to_Excel_Show(myWorkbook, , ".xlsx")

        Me.Cursor = Cursors.Default
        EntryMode = EntryMode_Save

        grdSOTTCLS1.ActiveRow = grdSOTTCLS1.Rows(0)
        Setup_SATBUDW1_TRADE_CLASS_CODE()
    End Sub

    Sub Create_Budget_Summary( _
        ByRef myWorkbook As GemBox.Spreadsheet.ExcelFile, _
        ByVal SUMMARY_CODE As String, _
        ByVal SUMMARY_DESC As String, _
        ByRef wsi As List(Of Integer))

        ASCMAIN1.Progress("Creating Summary")

        Dim cr As GemBox.Spreadsheet.CellRange

        Dim myWorksheet As GemBox.Spreadsheet.ExcelWorksheet = Nothing
        myWorksheet = myWorkbook.Worksheets.Add(SUMMARY_CODE & " " & "Summary")
        'myWorksheet = myWorkbook.Worksheets(SUMMARY_CODE & " " & "Summary")
        Dim wsn As String = "'" & "?" & "'!"
        Dim wsn2 As String = "'" & SUMMARY_CODE & " " & "Summary" & "'!"

        Dim T(10) As String

        Dim CTR As Integer = 0
        For Each i As Integer In wsi ' For i As Integer = 0 To myWorkbook.Worksheets.Count - 2

            Dim myWorksheet2 As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets(i)
            CTR += 1

            If CTR = 1 Then
                cr = myWorksheet2.Cells.GetSubrange("A1", "V6")
                cr.CopyTo(myWorksheet, "A1")
                myWorksheet.Cells("A3").Value = SUMMARY_CODE & " - " & SUMMARY_DESC

                For j As Integer = 0 To 21
                    myWorksheet.Columns(j).Width = myWorksheet2.Columns(j).Width
                Next

            End If

            For r2 As Integer = myWorksheet2.Rows.Count - 1 To 0 Step -1
                Dim XC As GemBox.Spreadsheet.ExcelCell = myWorksheet2.Cells(r2, 0)

                Dim rs As Integer = 0
                If CStr(XC.Value & "").StartsWith("Totals for") Then
                    cr = myWorksheet2.Cells.GetSubrange("A" & CStr(r2 + 1), "V" & CStr(r2 + 11))
                    rs = (CTR - 1) * 12 + 6 + 2
                    cr.CopyTo(myWorksheet, "A" & CStr(rs))

                    'For C As Integer = Asc("A") To Asc("V")
                    '    For j As Integer = rs + 1 To rs + 10
                    '        If myWorksheet.Cells(Chr(C) & CStr(j)).Formula & "" <> "" Then
                    '            '  Dim x As String = myWorksheet2.Name
                    '            myWorksheet.Cells(Chr(C) & CStr(j)).Formula = Nothing
                    '            myWorksheet.Cells(Chr(C) & CStr(j)).Value = myWorksheet2.Cells(Chr(C) & CStr(r2 + 1 + (j - rs))).Value
                    '        End If
                    '    Next
                    'Next
 

                    For j As Integer = 1 To 10
                        T(j) &= "+" & wsn & "A" & CStr(rs + j)
                    Next

                    If CTR = wsi.Count Then '  = myWorkbook.Worksheets.Count - 2 Then
                        rs = (CTR) * 12 + 6 + 2
                        cr.CopyTo(myWorksheet, "A" & CStr(rs))

                        myWorksheet.Cells("A" & CStr(rs)).Value = "Totals for " & SUMMARY_CODE & " - " & SUMMARY_DESC

                        For C As Integer = Asc("A") To Asc("V")
                            If Chr(C) < "B" Or Chr(C) > "C" Then

                                If Chr(C) = "D" Then
                                    T(7) = " 100 * If(" & wsn & "A" & CStr(rs + 3) & "=0,0," & wsn & "A" & CStr(rs + 6) & "/" & wsn & "A" & CStr(rs + 3) & ")"
                                    T(8) = " 100 * IF(" & wsn & "A" & CStr(rs + 4) & "=0,0,(+" & wsn & "A" & CStr(rs + 1) & "-" & wsn & "A" & CStr(rs + 4) & ")/" & wsn & "A" & CStr(rs + 4) & ")"
                                    T(9) = " 100 * IF(" & wsn & "A" & CStr(rs + 5) & "=0,0,(+" & wsn & "A" & CStr(rs + 1) & "-" & wsn & "A" & CStr(rs + 5) & ")/" & wsn & "A" & CStr(rs + 5) & ")"
                                End If

                                For j As Integer = 1 To 10
                                    myWorksheet.Cells(Chr(C) & CStr(rs + j)).Formula = "=" & Replace(Replace(Mid(T(j), 2), "A", Chr(C)), wsn, wsn2)
                                Next
                            End If
                        Next
                    End If
                    Exit For
                End If
            Next
        Next

        With myWorksheet.PrintOptions
            .BottomMargin = 0.25
            .TopMargin = 0.25
            .LeftMargin = 0.25
            .RightMargin = 0.25
            .FitWorksheetWidthToPages = 1
            .Portrait = False
        End With

        ASCMAIN1.Progress("", "")

    End Sub

    Sub Combined_Export()

        Dim myWorkbook As GemBox.Spreadsheet.ExcelFile = Nothing

        Dim r As Integer = 0
        Dim wsi As New List(Of Integer)

        ASCMAIN1.sql = "Select Distinct BRAND_CODE from SATBUDW1 where OPS_YYYY = '" & cmbOPS_YYYY.Value & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "BRAND_CODE")
            Dim BRAND_CODE As String = row.Item("BRAND_CODE")
            Absx1.txtFor("BRAND_CODE").Text = BRAND_CODE
            Click_Command("Load")
            Export_Budget_to_Excel(True, r, myWorkbook)
            wsi.Add(r - 1)
            Click_Command("Done")
        Next

        Create_Budget_Summary(myWorkbook, "All", "Brands", wsi)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Workbook")
        Gembox_Export_to_Excel_Show(myWorkbook, , ".xlsx")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
End Class