Imports Infragistics.Win.UltraWinGrid

Public Class SPFMKTG1
    Dim sqlSPTCOOPX As String
    Dim R_LAST As Integer = 0
    Dim COLS As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SPFSFOCI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("SPTPARM1")

        With dst
            sqlSPTCOOPX = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO" & vbCrLf _
            & ", SPTCOOP3.FEATURE_DESC, SPTCOOP3.ITEM_CODE, SPTTYPE1.SECURITY_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
            & " from SPTCOOP1,SPTCOOP3,ICTCOLL1,ICTBRAN1,SPTTYPE1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
            & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and SPTTYPE1.EXPENSE_TYPE_CODE (+) = SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
            & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO"
            ASCMAIN1.sql = sqlSPTCOOPX & "  and SPTCOOP1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False, "V")
            With .Tables("SPTCOOPX").Columns
                .Add("TOTAL_AMT", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("DIST_PCT", GetType(System.Decimal), "IIF(ISNULL(TOTAL_AMT,0)=0,0,100*ISNULL(DIST_AMT,0)/ISNULL(TOTAL_AMT,0))")
                .Add("DIST_OPEN", GetType(System.Decimal), "ISNULL(OPEN_AMT,0)*DIST_PCT/100")
                .Add("DIST_PAID", GetType(System.Decimal), "ISNULL(PAID_AMT,0)*DIST_PCT/100")
                .Add("DIST_OPEN_AND_PAID", GetType(System.Decimal), "ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0)")
                .Add("DIST_REAL_EXPENSE", GetType(System.Decimal), "IIF(OPS_YYYYPP>='" & Mid(ASCMAIN1.CYP, 1, 4) & "01" & "',ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0),IIF(ISNULL(DIST_PAID,0)=0,ISNULL(DIST_OPEN,0),ISNULL(DIST_PAID,0)))")
            End With
            .Tables("SPTCOOPX").PrimaryKey = New DataColumn() { .Tables("SPTCOOPX").Columns("AUTH_NO"), .Tables("SPTCOOPX").Columns("AUTH_LNO")}
            Create_TDA(.Tables.Add, "SPTMKTGG", "*")

            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)
            Fill_Records("ICTCOLL1")

            ASCMAIN1.sql = "Select * From GLTPARM3 Where YYYYMM Like :PARM1 || '%'"
            Create_TDA(.Tables.Add, "GLTPARMX", "**", 0, False, "V", 1)
            .Tables("GLTPARMX").Columns.Add("WEEK")


            'ASCMAIN1.sql = "Select * From GLTPARM3 Where YYYYMM Like '2023%'"
            'Create_TDA(.Tables.Add, "GLTPARMX", "**", 0, False, "", 1)
            '.Tables("GLTPARMX").Columns.Add("WEEK")
            'Fill_Records("GLTPARMX")
            'Dim W As Integer = 0
            'For Each ROW As DataRow In dst.Tables("GLTPARMX").Select("", "YYYYWW")
            '    W += 1
            '    ROW.Item("WEEK") = "Week " & CStr(W)
            'Next

            'ASCMAIN1.sql = "Select * from ASTCODE1 WHERE TABLE_NAME = 'SPTCOOP1' AND COLUMN_NAME = 'APPR_STATUS_CODE'"
            Create_TDA(.Tables.Add, "ASTCODE1", "*", 2, False)
            Fill_Records("ASTCODE1", New String() {"SPTCOOP1", "APPR_STATUS_CODE"})
        End With

        grdSPTCOOPX.DataSource = dst.Tables("SPTCOOPX")
        grdSPTMKTGG.DataSource = dst.Tables("SPTMKTGG")

        Create_Summary(grdSPTCOOPX, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOPX, New String() {"DIST_AMT", "OPEN_AMT", "PAID_AMT", "DIST_OPEN", "DIST_PAID", "DIST_OPEN_AND_PAID", "DIST_REAL_EXPENSE"})

        With grdSPTCOOPX.DisplayLayout.Bands("SPTCOOPX")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"NOTES", "BOOKING_NAME", "VERIFIED_AS_OPEN_COMMENTS"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If New String() {"OPEN_AMT", "PAID_AMT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"COLLECTION_CODE", "DIST_AMT", "DIST_PCT", "DIST_OPEN", "DIST_PAID", "DIST_OPEN_AND_PAID"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"AUTH_APPR_DATE", "AUTH_APPR_BY", "AUTH_APPR_AMT", "AUTH_APPR_NOTES"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                ElseIf New String() {"VERIFIED_AS_OPEN", "VERIFIED_AS_OPEN_NOTES", "VERIFIED_AS_OPEN_AMT", "VERIFIED_AS_OPEN_BY", "VERIFIED_AS_OPEN_DATE", "VERIFIED_AS_OPEN_COMMENTS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTMKTGG.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen

                GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit

                If New String() {"EXPENSE_TYPE_CODE", "VEHICLE_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"SPEND_TYPE", "BOOKED_BY", "ASSETS_USED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"MEDIA_CATGY", "MEDIA_TYPE", "MEDIA_DESC"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                ElseIf New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 2 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("BUDGET_YEAR").DataSource = YEARs ' New String() {"2008", "2009", "2010"}
        Absx1.cbeFor("BUDGET_YEAR").Value = Now.Year


        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "BOOKED_BY", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'BOOKED_BY'")

        Show_Filter(grdSPTCOOPX, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

            Case "Export"
                Dim p1 As Integer = trkFrom.Value
                Dim p2 As Integer = trkTo.Value

                'declare 2 more maybe, pf and pt
                If p1 > p2 Then
                    EMsg &= vbCr & "Warning: Start quarter cannot be greater than end quarter."
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Export"
                Export()
                SplitContainer1.Panel2Collapsed = False

            Case "Done", "Cancel"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")

                    'If Not ScreenMode Then
                    '    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    'Else
                    '    .Items("Edit").Settings.Enabled = not_iScreenMode
                    'End If

                    .Items("Edit").Visible = Not ScreenMode
                    .Items("View").Visible = Not ScreenMode

                    '.Items("View").Settings.Enabled = not_iScreenMode

                    'If ScreenMode And EntryMode <> "V" Then
                    '    .Items("Update").Settings.Enabled = not_iScreenMode
                    '    .Items("Cancel").Settings.Enabled = not_iScreenMode
                    'Else
                    '    .Items("Update").Settings.Enabled = iScreenMode
                    '    .Items("Cancel").Settings.Enabled = iScreenMode
                    'End If

                    .Items("Update").Visible = ScreenMode And (EntryMode = "E")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "E")

                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")

                    .Items("Export").Visible = ScreenMode And (EntryMode = "V")

                    'If ScreenMode And EntryMode <> "V" Then
                    '    .Items("Done").Settings.Enabled = not_iScreenMode
                    'Else
                    '    .Items("Done").Settings.Enabled = iScreenMode
                    'End If
                End With
                .Groups("Column Set").Visible = False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode And (EntryMode = "V")
        grpHeader.Visible = ScreenMode

        grdSPTMKTGG.Visible = Not ScreenMode Or (EntryMode = "E")
        grdSPTCOOPX.Visible = ScreenMode And (EntryMode = "V")

        If ScreenMode Then
            'Set_Read_Only(grpHeader, (EntryMode = "V"))
            If EntryMode = "V" Then
                SplitContainer1.Panel2Collapsed = True
            ElseIf EntryMode = "E" Then
                With grdSPTMKTGG.DisplayLayout.Override
                    .AllowAddNew = AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            End If
        Else
            Clear_Record()
            grdSPTCOOPX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

            With grdSPTMKTGG.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SPTCOOPX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "V" Then
            Dim YYYY As String = cbeBUDGET_YEAR.Value

            Fill_Records("GLTPARMX", YYYY)
            Dim W As Integer = 0
            For Each ROW As DataRow In dst.Tables("GLTPARMX").Select("", "YYYYWW")
                W += 1
                ROW.Item("WEEK") = "Week " & CStr(W)
            Next

            Dim p1 As Integer = trkFrom.Value
            Dim p2 As Integer = trkTo.Value
            Dim pf As String = ""
            Dim pt As String = ""
            'declare 2 more maybe, pf and pt
            pf = IIf(p1 = 1, "01", IIf(p1 = 2, "04", IIf(p1 = 3, "07", IIf(p1 = 4, "10", "00"))))
            pt = Format((p2 - 1) * 3 + 3, "00")

            ASCMAIN1.sql = sqlSPTCOOPX & vbCrLf _
                & $" and SPTCOOP1.OPS_YYYYPP between '{YYYY}{pf}' and '{YYYY}{pt}' " & vbCrLf _
                & " and (SPTCOOP1.EXPENSE_TYPE_CODE, SPTCOOP1.VEHICLE_CODE) in (Select EXPENSE_TYPE_CODE, VEHICLE_CODE from SPTMKTGG)"

            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSPTCOOPX, "AUTH_NO".ToLower)
            'grdSPTCOOPX.Text = "All"
        End If

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Update_Record_TDA("SPTMKTGG")
        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        ' Dependent_Updates(-1, ORDR_NO)
        'For Each TABLE_NAME As String In New String() _
        '    {"SPTSFOC1", "SPTSFOC3", "SPTSFOC9"}
        '    Delete_Records_1(TABLE_NAME)
        'Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("EVENT_GROUP_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "VEHICLE_CODE"
                sql_where = "VEHICLE_CODE in ('BF','MA')"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTCOOPX, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name

            'Case "grdSPTSFOC9"
            '    tlb_btn = tlb_pop.Tools("Load All Stores")
            '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
            '        tlb_btn.SharedProps.Visible = True
            '    Else
            '        tlb_btn.SharedProps.Visible = False
            '    End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdSPTSFOC9"
                '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                '        tlb_btn.SharedProps.Visible = True
                '    Else
                '        tlb_btn.SharedProps.Visible = False
                '    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            'Case "Select All", "De-Select All"
            '    Me.Cursor = Cursors.WaitCursor
            '    ASCMAIN1.Progress("Now executing " & e.Tool.Key)
            '    For Each grow As UltraWinGrid.UltraGridRow In grdSPTSFOC9.Rows
            '        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
            '        grow.Update()
            '    Next
            '    Me.Cursor = Cursors.Default
            '    ASCMAIN1.Progress("")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            'Case "Item Status Inquiry"
            '    Dim VEHICLE_CODE As String = grd.ActiveRow.Cells("VEHICLE_CODE").Text
            '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEHICLE_CODE)
            '    If rowSPTAVEH1 IsNot Nothing Then
            '        Context_Launch("View", VEHICLE_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "EVENT_GROUP_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEHICLE_CODE"
                Dim VEHICLE_CODE As String = Absx1.txtFor("VEHICLE_CODE").Text
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "EVENT_GROUP_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "TOTAL_AMT"
                ' Calculate_OPEN_AMT()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "APPR_STATUS_CODE"
                If Absx1.optFor("APPR_STATUS_CODE").Value = "X" Then
                    Absx1.optFor("STATUS_CODE").Value = "C"
                Else

                End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "AUTH_DATE"
                If Absx1.dteFor("AUTH_DATE").Value & "" = "" Then
                    Absx1.txtFor("OPS_YYYYWW").Text = ""
                Else
                    Dim DATE_START As Date = Absx1.dteFor("AUTH_DATE").Value
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
                        Dim YW As String = ASCDATA1.GetDataValue
                        If YW <> "" Then
                            Absx1.txtFor("OPS_YYYYWW").Text = YW
                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Documents")

        EnforceConstraints(False)
        If ScreenMode Then
            ASCMAIN1.sql = sqlSPTCOOPX
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSPTCOOPX, "AUTH_NO".ToLower)
            grdSPTCOOPX.Text = "All"
        Else
            ASCMAIN1.sql = "Select * from SPTMKTGG"
            Fill_Records("SPTMKTGG", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSPTMKTGG, "EXPENSE_TYPE_CODE,VEHICLE_CODE")
            grdSPTMKTGG.Text = "Expense Types / Vehicles to pull from Promo Events Database"
        End If

        EnforceConstraints(True)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")
    End Sub

#Region "grdSPTMKTGG"
    Private Sub grdSPTMKTGG_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSPTMKTGG.ClickCellButton

        If grdSPTMKTGG.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "EXPENSE_TYPE_CODE"
                'sql_where = "ITEM_STATUS = 'A'"
            Case "VEHICLE_CODE"
                'sql_where = "ITEM_STATUS = 'A'"
        End Select
        grdClickCellButton(grdSPTMKTGG, sql_where, False)

    End Sub

    Private Sub grdSPTCOOPX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSPTCOOPX.InitializeLayout

    End Sub


    Sub Export()
        WorkbookView1.GetLock()
        'Dim tbl As DataTable = dst.Tables("ARTATBR1").Copy
        'For Each row As DataRow In dst.Tables("ARTATBR1").Select
        '    tbl.Rows.Add(row.ItemArray)
        'Next
        Dim APPR_STATUS_CODEs As New Dictionary(Of String, String)
        For Each row As DataRow In dst.Tables("ASTCODE1").Select()
            Dim APPR_STATUS_CODE As String = row.Item("T_CODE")
            Dim APPR_STATUS_DESC As String = row.Item("T_DESC")
            APPR_STATUS_CODEs.Add(APPR_STATUS_CODE, APPR_STATUS_DESC)
        Next

        ' Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        oSheet = WorkbookView1.ActiveWorkbook.Worksheets(0)
        oSheet.Cells.Clear()



        oSheet.Cells(0, 0, 0, 24).EntireColumn.NumberFormat = "@"
        oSheet.Cells(0, 14).EntireColumn.NumberFormat = "#,##0"
        oSheet.Cells(0, 21).EntireColumn.NumberFormat = "#,##0"
        oSheet.Cells(0, 23).EntireColumn.NumberFormat = "#,##0"

        oSheet.Cells(0, 0).EntireRow.NumberFormat = "@"
        oSheet.Cells(0, 0).EntireRow.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        oSheet.Cells(0, 0).EntireRow.Interior.Color = SpreadsheetGear.Colors.LightBlue

        oSheet.Cells(0, 1).Value = "Brand Line"
        oSheet.Cells(0, 3).Value = "Type"
        oSheet.Cells(0, 4).Value = "Booked By"
        oSheet.Cells(0, 5).Value = "Assets Used"
        oSheet.Cells(0, 7).Value = "Media Category"
        oSheet.Cells(0, 8).Value = "Media Type"
        oSheet.Cells(0, 9).Value = "Media Desc"
        oSheet.Cells(0, 10).Value = "Program Details"
        oSheet.Cells(0, 12).Value = "Start"
        oSheet.Cells(0, 13).Value = "Year"
        oSheet.Cells(0, 14).Value = "Weeks"
        oSheet.Cells(0, 20).Value = "Impressions"
        oSheet.Cells(0, 21).Value = "Units"
        oSheet.Cells(0, 23).Value = "Budget$"

        Dim R As Integer = 0
        For Each grow As UltraWinGrid.UltraGridRow In grdSPTCOOPX.Rows
            If grow.IsFilteredOut Then
            Else
                Dim OPS_YYYYPP As String = grow.Cells("OPS_YYYYPP").Value & ""
                Dim AUTH_NO As String = (grow.Cells("AUTH_NO").Value & "")
                Dim AUTH_LNO As String = Val(grow.Cells("AUTH_LNO").Value & "")
                Dim row As DataRow = dst.Tables("SPTCOOPX").Rows.Find(New Object() {AUTH_NO, AUTH_LNO})


                Dim EXPENSE_TYPE_CODE As String = row.Item("EXPENSE_TYPE_CODE") & "" '1
                    Dim VEHICLE_CODE As String = row.Item("VEHICLE_CODE") & "" '2
                    Dim rowSPTMKTGG As DataRow = dst.Tables("SPTMKTGG").Rows.Find(New String() {EXPENSE_TYPE_CODE, VEHICLE_CODE})
                    Dim OPS_YYYYWW As String = row.Item("OPS_YYYYWW") & ""
                    Dim rowGLTPARMX As DataRow = dst.Tables("GLTPARMX").Rows.Find(OPS_YYYYWW)
                    Dim WEEK As String = "Week"
                    If rowGLTPARMX IsNot Nothing Then
                        WEEK = rowGLTPARMX.Item("WEEK") & ""
                    End If

                    Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                    Dim COLLECTION_NAME_MKTG As String = ""
                    If rowICTCOLL1 Is Nothing Then
                        COLLECTION_NAME_MKTG = COLLECTION_CODE
                    Else
                        COLLECTION_NAME_MKTG = rowICTCOLL1.Item("COLLECTION_NAME_MKTG") & ""
                        If COLLECTION_NAME_MKTG = "" Then
                            COLLECTION_NAME_MKTG = rowICTCOLL1.Item("COLLECTION_NAME") & ""
                            COLLECTION_NAME_MKTG = COLLECTION_NAME_MKTG.ToUpper
                        End If
                    End If

                    Dim APPR_STATUS_CODE As String = row.Item("APPR_STATUS_CODE")
                    Dim APPR_STATUS_DESC As String = APPR_STATUS_CODE
                    If APPR_STATUS_CODEs.ContainsKey(APPR_STATUS_CODE) Then
                        APPR_STATUS_DESC = APPR_STATUS_CODEs(APPR_STATUS_CODE)
                    End If

                    Dim QTY As Int32 = Val(row.Item("QTY") & "")

                    Dim DIST_OPEN_AND_PAID As Decimal = Val(row.Item("DIST_OPEN_AND_PAID") & "")

                    Dim DATE_START As Date = CDate(row.Item("DATE_START"))
                    Dim DATE_END As Date = CDate(row.Item("DATE_END"))
                    Dim DAYS As Integer = DATE_END.Subtract(DATE_START).TotalDays
                    If (DAYS Mod 7) <> 0 Then
                        DAYS = DAYS + 7 - (DAYS Mod 7)
                    End If
                    Dim WKS As Integer = DAYS / 7
                    If WKS = 0 Then WKS = 1

                    Dim W As Integer = 1

                    Dim C As Integer = -1
                    R += 1

                    C += 1 : oSheet.Cells(R, C).Value = row.Item("AUTH_NO") '0
                    C += 1 : oSheet.Cells(R, C).Value = COLLECTION_NAME_MKTG ' row.Item("BRAND_CODE") '1
                    C += 1 : oSheet.Cells(R, C).Value = "" '2
                    C += 1 : oSheet.Cells(R, C).Value = rowSPTMKTGG.Item("SPEND_TYPE") '3
                    Dim BOOKED_BY As String = rowSPTMKTGG.Item("BOOKED_BY") & ""
                    If row.Item("BOOKED_BY") & "" <> "" Then
                        BOOKED_BY = row.Item("BOOKED_BY")
                    End If
                    C += 1 : oSheet.Cells(R, C).Value = BOOKED_BY '4
                    C += 1 : oSheet.Cells(R, C).Value = rowSPTMKTGG.Item("ASSETS_USED") ' 5
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 6
                    C += 1 : oSheet.Cells(R, C).Value = rowSPTMKTGG.Item("MEDIA_CATGY") ' 7
                    C += 1 : oSheet.Cells(R, C).Value = rowSPTMKTGG.Item("MEDIA_TYPE") ' 8
                    'rowSPTMKTGG.Item("MEDIA_DESC") <-- HAVE DESC = THIS IF WE DONT WANT BLANKS
                    Dim desc As String = ""
                    If row.Item("CUST_CODE") = "AMAZON" Then
                        desc = "AMAZON"
                    End If
                    C += 1 : oSheet.Cells(R, C).Value = desc ' 9 Media desc 
                    C += 1 : oSheet.Cells(R, C).Value = row.Item("CUST_CODE") & " " & row.Item("BOOKING_NAME") ' 10
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 11
                    C += 1 : oSheet.Cells(R, C).Value = WEEK ' 12
                    C += 1 : oSheet.Cells(R, C).Value = cbeBUDGET_YEAR.Value ' 13
                    C += 1 : oSheet.Cells(R, C).Value = CStr(WKS) ' 14
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 15
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 16
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 17
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 18
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 19
                    C += 1 : oSheet.Cells(R, C).Value = "Impressions" ' 20
                    C += 1 : oSheet.Cells(R, C).Value = IIf(QTY = 0, "", Format(QTY, "#.##0")) ' 21
                    C += 1 : oSheet.Cells(R, C).Value = "" ' 22
                    C += 1 : oSheet.Cells(R, C).Value = Format(DIST_OPEN_AND_PAID, "#,##0") ' 23
                    C += 1 : oSheet.Cells(R, C).Value = APPR_STATUS_DESC ' 24



            End If
        Next

        R_LAST = R

        oSheet.Cells(0, 0, R_LAST, 23).EntireColumn.AutoFit()

        COLS.Clear()
        COLS.Add("PRODUCT", $"{Excel_Cell0(1, 1)}:{Excel_Cell0(R_LAST, 1)}")
        COLS.Add("BUDGET DETAILS", $"{Excel_Cell0(1, 3)}:{Excel_Cell0(R_LAST, 5)}")
        COLS.Add("CAMPAIGN TYPE", $"{Excel_Cell0(1, 7)}:{Excel_Cell0(R_LAST, 10)}")
        COLS.Add("PLANNING", $"{Excel_Cell0(1, 12)}:{Excel_Cell0(R_LAST, 14)}")
        COLS.Add("KPIs", $"{Excel_Cell0(1, 20)}:{Excel_Cell0(R_LAST, 21)}")
        COLS.Add("BUDGET", $"{Excel_Cell0(1, 23)}:{Excel_Cell0(R_LAST, 23)}")

        'Load_DataTable_into_SGXLS(1, 1, dst.Tables("SPTMKTGE"), WorkbookView1.ActiveWorksheet, grdSPTMKTGG, Nothing, "", "")
        WorkbookView1.ReleaseLock()

        UltraExplorerBar1.Groups("Column Set").Visible = True
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & ASCMAIN1.Next_Control_No($"{Me.Name}.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub grdSPTMKTGG_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSPTMKTGG.BeforeRowUpdate

        Dim row As DataRow

        Dim EXPENSE_TYPE_CODE As String = e.Row.Cells("EXPENSE_TYPE_CODE").Value & ""
        Dim VEHICLE_CODE As String = e.Row.Cells("VEHICLE_CODE").Value & ""

        Dim EMsg As String = ""

        row = LookUp("SPTTYPE1", EXPENSE_TYPE_CODE)
        If row Is Nothing Then
            EMsg &= vbCr & "Invalid Value for Expense Type Code"
            e.Cancel = True
        End If

        row = LookUp("SPTAVEH1", VEHICLE_CODE)
        If row Is Nothing Then
            EMsg &= vbCr & "Invalid Value for Vehicle Code"
            e.Cancel = True
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update this Record")
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
        End If
        e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
    End Sub

    Private Sub cmdCopy2Clipboard_Click(sender As Object, e As EventArgs) Handles cmdCopy2Clipboard.Click
        Dim XLR As String = COLS(optCols.Text)
        WorkbookView1.GetLock()
        WorkbookView1.ActiveWorkbook.Worksheets(0).Range(XLR).Select()
        WorkbookView1.Copy()
        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub trkFrom_ValueChanged(sender As Object, e As EventArgs) Handles trkFrom.ValueChanged
        Set_Q_Labels()
    End Sub
    Private Sub trkTo_ValueChanged(sender As Object, e As EventArgs) Handles trkTo.ValueChanged
        Set_Q_Labels()
    End Sub

    Sub Set_Q_Labels()
        lblQFrom.Text = "Q" & CStr(trkFrom.Value)
        lblQTo.Text = "Q" & CStr(trkTo.Value)
    End Sub

    Private Sub UltraLabel1_Click(sender As Object, e As EventArgs) Handles UltraLabel1.Click

    End Sub
#End Region
End Class