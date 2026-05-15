Imports Infragistics.Win.UltraWinGrid
Public Class WHFCALL1
    Dim SHIP_TO_TYPE As String
    Dim SHIP_TO_KEY As String
    Dim SHIP_TO_CODE_PFX As String
    Dim sqlSOTSHIPI As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SOTSHIPI.*" & vbCrLf _
                & ", DECODE(SOTSHIPD.SHIP_TO_TYPE,'C',ARTCUST1.CUST_NAME,'W',ICTWHSE1.WHSE_DESC,'S',ARTCUST2.CUST_STORE_NAME,'V',APTVEND1.VEND_NAME,'?') SHIP_TO_NAME" & vbCrLf _
                & " from SOTSHIPI,SOTSHIPD,SOTSHIPE,ARTCUST1,ARTCUST2,APTVEND1,ICTWHSE1" & vbCrLf _
                & " where NVL(SOTSHIPI.CALL_IN_RULE_STATUS,'?') = :PARM1" & vbCrLf _
                & "   and (:PARM2 = '*' or SOTSHIPE.SHIP_TO_CODE_PFX = :PARM2)" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE (+) = SOTSHIPI.SHIP_TO_CODE" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE_PFX = SOTSHIPE.SHIP_TO_CODE_PFX" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO (+) = SOTSHIPD.CUST_STORE_NO" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = SOTSHIPD.VEND_CODE" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE (+) = SOTSHIPD.WHSE_CODE"
            sqlSOTSHIPI = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTSHIPI", "**", 0, True, "VV", 1)

            ASCMAIN1.sql = sqlSOTSHIPI
            Create_TDA(.Tables.Add, "SOTSHIPI_INACTIVE", "**", 0, False, "VV", 1)

            For Each T As String In New String() {"SOTSHIPI", "SOTSHIPI_INACTIVE"}
                With dst.Tables(T)
                    For Each D As String In New String() {"SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"}
                        .Columns.Add($"CALL_IN_{D}")
                        .Columns($"CALL_IN_{D}").DefaultValue = "0"
                        .Columns.Add($"PICK_UP_{D}")
                        .Columns($"PICK_UP_{D}").DefaultValue = "0"
                    Next
                End With
            Next

            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE" & vbCrLf _
                & ", ARTCUST2.CUST_DC_IND, ARTCUST2.CUST_DC_NO, SOTSHIPD.SHIP_TO_CODE" & vbCrLf _
                & " from ARTCUST2, SOTSHIPD" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = :PARM1" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SOTSHIPD.CUST_CODE (+)" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SOTSHIPD.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "ARTCUST2_SHIPTO", "**", 0, False, "V", 2)
            With dst.Tables("ARTCUST2_SHIPTO")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

        End With

        grdSOTSHIPI.DataSource = dst.Tables("SOTSHIPI")
        Create_Summary(grdSOTSHIPI, "CALL_IN_RULE_NO", "Count")
        Show_Filter(grdSOTSHIPI)


        grdSOTSHIPI_INACTIVE.DataSource = dst.Tables("SOTSHIPI_INACTIVE")
        Create_Summary(grdSOTSHIPI_INACTIVE, "CALL_IN_RULE_NO", "Count")
        Show_Filter(grdSOTSHIPI_INACTIVE)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTSHIPI, grdSOTSHIPI_INACTIVE}

            With grd.DisplayLayout.Bands(0)

                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.CellActivation = Activation.NoEdit
                    If grd.Name = "grdSOTSHIPI" Then
                        With grdSOTSHIPI_INACTIVE.DisplayLayout.Bands(0).Columns(gcol.Key)
                            .Header.Caption = gcol.Header.Caption
                            .Width = gcol.Width
                            .Style = gcol.Style
                            .Format = gcol.Format
                            .CellAppearance.TextHAlign = gcol.CellAppearance.TextHAlign
                            .Header.Appearance.TextHAlign = gcol.Header.Appearance.TextHAlign
                        End With
                    End If
                Next

                Dim g As New UltraWinGrid.UltraGridGroup

                g = .Groups.Add("RECORD_KEYS")
                g.Header.Fixed = True
                g.Header.Caption = "Record Keys"
                For Each c As String In New String() {"CALL_IN_RULE_NO", "CALL_IN_RULE_STATUS", "SHIP_TO_CODE", "SHIP_TO_NAME"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Fixed = True
                Next

                g = .Groups.Add("CALL_IN_RULES")
                g.Header.Caption = "Call-In Rules"
                For Each c As String In New String() {"CALL_IN_NOTICE_DAYS"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Yellow
                    If grd.Name = "grdSOTSHIPI" Then .Columns(c).CellActivation = Activation.AllowEdit
                Next

                g = .Groups.Add("CALL_IN_DAYS")
                g.Header.Caption = "Call-In Days"
                For Each c As String In New String() {"CALL_IN_HOL_EXC", "CALL_IN_SUN", "CALL_IN_MON", "CALL_IN_TUE", "CALL_IN_WED", "CALL_IN_THU", "CALL_IN_FRI", "CALL_IN_SAT"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    If grd.Name = "grdSOTSHIPI" Then .Columns(c).CellActivation = Activation.AllowEdit
                Next

                g = .Groups.Add("PICK_UP_DAYS")
                g.Header.Caption = "Pick-Up Days"
                For Each c As String In New String() {"PICK_UP_HOL_EXC", "PICK_UP_SUN", "PICK_UP_MON", "PICK_UP_TUE", "PICK_UP_WED", "PICK_UP_THU", "PICK_UP_FRI", "PICK_UP_SAT"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    If grd.Name = "grdSOTSHIPI" Then .Columns(c).CellActivation = Activation.AllowEdit
                Next

                g = .Groups.Add("AS400 Key")
                g.Header.Caption = "AS400 Ship-To"
                For Each c As String In New String() {"SHIPTO1", "SHIPTO2"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Next

                g = .Groups.Add("CHANGES")
                g.Header.Caption = "Dates Changed"
                For Each c As String In New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                Next

                For Each grp As UltraWinGrid.UltraGridGroup In .Groups
                    grp.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    grp.Header.Appearance.BackColor = System.Drawing.Color.White
                    grp.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Next

                ASCMAIN1.Add_Value_List(grd, "CALL_IN_HOL_EXC", Nothing, New String() {":", "E:1 Bus Day Before", "S:Next Sched Day", "L:1 Bus Day Later"})
                ASCMAIN1.Add_Value_List(grd, "PICK_UP_HOL_EXC", Nothing, New String() {":", "E:1 Bus Day Before", "S:Next Sched Day", "L:1 Bus Day Later"})

                ASCMAIN1.Add_Value_List(grd, "CALL_IN_RULE_STATUS", Nothing, New String() {":", "A:Active", "I:Retired"})

            End With
        Next


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                SHIP_TO_KEY = Absx1.txtFor("SHIP_TO_KEY").Text
                SHIP_TO_TYPE = optSHIP_TO_TYPE.Value
                SHIP_TO_CODE_PFX = Absx1.txtFor("SHIP_TO_CODE_PFX").Text

                Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)
                If rowSOTSHIPE Is Nothing Then
                    EMsg &= vbCr & "Invalid Ship-To Code Prefix"
                End If

                If EMsg = "" Then
                    ASCMAIN1.Logical_Lock("SOTSHIPE", SHIP_TO_CODE_PFX)
                End If

            Case "Validate All"
                EMsg = Validate_Active_Records()
                If EMsg = "" Then
                    MsgBox("All Call-In Records are Valid", MsgBoxStyle.OkOnly, "Validation Complete")
                Else
                    Using frm As New ASFMSGBF
                        EMsg = Replace(EMsg, vbCr, "</br>")
                        frm.Show_Formatted_txt("Data Validation Issues", EMsg, Me)
                        EMsg = ""
                    End Using
                End If

            Case "Done"

            Case "Update"

                EMsg = Validate_Active_Records

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
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Validate All").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Visible = tf
            .Groups("Screen Control").Items("Update").Visible = tf
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdSOTSHIPI.Visible = ScreenMode
        'grdSOTSHIPI_INACTIVE.Visible = ScreenMode

        grpLead.Visible = Not ScreenMode
        SplitContainer1.Visible = ScreenMode

        If Not ScreenMode Then
            Clear_Record()
            grdSOTSHIPI.Parent = grpLead
            grdSOTSHIPI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else

            grdSOTSHIPI.Parent = SplitContainer1.Panel1
            grdSOTSHIPI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)
            'Dim SHIP_TO_CODE_PFX As String = rowSOTSHIPE.Item("SHIP_TO_CODE_PFX") & ""
            'Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX

            Select Case SHIP_TO_TYPE
                Case "C"
                    'lblSHIP_TO_TYPE.Text = "Customer"
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowSOTSHIPE.Item("CUST_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowARTCUST1.Item("CUST_NAME")
                'Case "S"
                '    lblSHIP_TO_TYPE.Text = "DC/Store"
                Case "W"
                    'lblSHIP_TO_TYPE.Text = "Warehouse"
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIPE.Item("WHSE_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowICTWHSE1.Item("WHSE_DESC")
                Case "V"
                    'lblSHIP_TO_TYPE.Text = "Vendor"
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1 ", rowSOTSHIPE.Item("VEND_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowAPTVEND1.Item("VEND_NAME")
            End Select
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPI", "SOTSHIPI_INACTIVE"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("SHIP_TO_KEY").Text = ""
        Absx1.txtFor("SHIP_TO_KEY_NAME").Text = ""
        Absx1.txtFor("SHIP_TO_CODE_PFX").Text = ""

        Refresh_Records()

        Clear_All_Filters(grdSOTSHIPI)
        Clear_All_Filters(grdSOTSHIPI_INACTIVE)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Fill_Records("SOTSHIPI", New String() {"A", SHIP_TO_CODE_PFX})

        Fill_Records("SOTSHIPI_INACTIVE", New String() {"I", SHIP_TO_CODE_PFX})

        Fix_CheckBoxes

        dst.AcceptChanges()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Updating...")

        For Each rowSOTSHIPI As DataRow In dst.Tables("SOTSHIPI").Select("CALL_IN_RULE_STATUS = 'A'", "", DataViewRowState.ModifiedCurrent)
            Dim CALL_IN_RULE_NO As String = rowSOTSHIPI.Item("CALL_IN_RULE_NO")

            Dim rowSOTSHIPI_CHANGED As DataRow = dst.Tables("SOTSHIPI").NewRow
            rowSOTSHIPI_CHANGED.ItemArray = rowSOTSHIPI.ItemArray
            rowSOTSHIPI_CHANGED.Item("CALL_IN_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPI.CALL_IN_RULE_NO")
            dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI_CHANGED)

            Dim rowSOTSHIPI_ORIG As DataRow = dst.Tables("SOTSHIPI").Select($"CALL_IN_RULE_NO = '{CALL_IN_RULE_NO}'", "", DataViewRowState.OriginalRows)(0)
            rowSOTSHIPI_ORIG.RejectChanges()
            rowSOTSHIPI_ORIG.Item("CALL_IN_RULE_STATUS") = "I"
        Next

        For Each rowSOTSHIPI_INACTIVE As DataRow In dst.Tables("SOTSHIPI_INACTIVE").Select
            Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").NewRow
            rowSOTSHIPI.ItemArray = rowSOTSHIPI_INACTIVE.ItemArray
            dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI)
        Next

        Update_Record_TDA("SOTSHIPI", $"SHIP_TO_CODE like '{SHIP_TO_CODE_PFX}%'")

        MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Verification")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_TO_KEY"
                sql_where &= $"SHIP_TO_TYPE = '{optSHIP_TO_TYPE.Value}'"

            Case "SHIP_TO_CODE_PFX"
                sql_where &= $"SHIP_TO_TYPE = '{optSHIP_TO_TYPE.Value}'"

            Case "SHIP_TO_CODE"
                sql_where &= "CUST_STORE_NO IS NULL"

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPI, "SSBBB", "Show Filter", "Show GroupBox", "Retire", "Create Default Rule for Key", "Copy Rule to Key", "Create Default Rule for New Ship-Tos", "Copy Rule to New Ship-Tos")
        Load_Popup_Menu(grdSOTSHIPI_INACTIVE, "SSB", "Show Filter", "Show GroupBox", "Re-Activate")
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
            Select Case e.SourceControl.Name

                Case "grdSOTSHIPI"
                    tlb_btn = DirectCast(tlb_pop.Tools("Retire"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Default Rule for Key"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
                    If tlb_btn.SharedProps.Visible Then
                        tlb_btn.SharedProps.Caption = $"Create Default Rule for {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}"
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Rule to Key"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
                    If tlb_btn.SharedProps.Visible Then
                        tlb_btn.SharedProps.Caption = $"Copy Rule to {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}"
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Default Rule for New Ship-Tos"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E" AndAlso optSHIP_TO_TYPE.Value = "C"

                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Rule to New Ship-Tos"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E" AndAlso optSHIP_TO_TYPE.Value = "C"

                Case "grdSOTSHIPI_INACTIVE"
                    tlb_btn = DirectCast(tlb_pop.Tools("Re-Activate"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case ""

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Create Default Rule for Key"
                Dim rowKey() As DataRow = dst.Tables("SOTSHIPI").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE_PFX}D'")

                Dim CALL_IN_RULE_NO As String = ""
                If rowKey.Length > 0 Then
                    CALL_IN_RULE_NO = rowKey(0).Item("CALL_IN_RULE_NO")
                    MsgBox($"There is already a Call-In Rule ({CALL_IN_RULE_NO})" & vbCrLf & $" set up for the {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                Dim SHIP_TO_CODE As String = SHIP_TO_CODE_PFX & "D"
                Dim SHIP_TO_NAME As String = Absx1.txtFor("SHIP_TO_KEY_NAME").Text
                Create_Default_CALL_IN_rule(SHIP_TO_CODE, SHIP_TO_NAME)

                Sort_grdColumns(grdSOTSHIPI, "SHIP_TO_CODE")

            Case "Copy Rule to Key"
                Dim rowKey() As DataRow = dst.Tables("SOTSHIPI").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE_PFX}D'")

                Dim CALL_IN_RULE_NO As String = ""
                If rowKey.Length > 0 Then
                    CALL_IN_RULE_NO = rowKey(0).Item("CALL_IN_RULE_NO")
                    MsgBox($"There is already a Call-In Rule ({CALL_IN_RULE_NO})" & vbCrLf & $" set up for the {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                If grdSOTSHIPI.Selected.Rows.Count = 0 And grdSOTSHIPI.ActiveRow IsNot Nothing Then
                    grdSOTSHIPI.ActiveRow.Selected = True
                End If

                If grdSOTSHIPI.Selected.Rows.Count > 1 Then
                    MsgBox($"You may select only 1 row to copy to {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                CALL_IN_RULE_NO = grdSOTSHIPI.Selected.Rows(0).Cells("CALL_IN_RULE_NO").Value
                Dim rowSOTSHIPI_RULE As DataRow = dst.Tables("SOTSHIPI").Rows.Find(CALL_IN_RULE_NO)

                Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").NewRow
                rowSOTSHIPI.ItemArray = rowSOTSHIPI_RULE.ItemArray
                rowSOTSHIPI.Item("CALL_IN_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPI.CALL_IN_RULE_NO")
                rowSOTSHIPI.Item("SHIP_TO_CODE") = SHIP_TO_CODE_PFX & "D"
                rowSOTSHIPI.Item("SHIP_TO_NAME") = Absx1.txtFor("SHIP_TO_KEY_NAME").Text
                rowSOTSHIPI.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPI.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPI.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTSHIPI.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIPI.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPI.Item("LAST_OPER") = ASCMAIN1.USER_ID

                dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI)
                Sort_grdColumns(grdSOTSHIPI, "SHIP_TO_CODE")

            Case "Retire"

                If grdSOTSHIPI.Selected.Rows.Count = 0 And grdSOTSHIPI.ActiveRow IsNot Nothing Then
                    grdSOTSHIPI.ActiveRow.Selected = True
                End If

                If grdSOTSHIPI.Selected.Rows.Count > 0 Then
                    If MsgBox($"OK to Retired the {grdSOTSHIPI.Selected.Rows.Count} selected row(s)?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Dim CALL_IN_RULE_NOs As New List(Of String)
                        For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPI.Selected.Rows
                            Dim CALL_IN_RULE_NO As String = grow.Cells("CALL_IN_RULE_NO").Value
                            CALL_IN_RULE_NOs.Add(CALL_IN_RULE_NO)
                        Next
                        For Each CALL_IN_RULE_NO As String In CALL_IN_RULE_NOs
                            'Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").Rows.Find(CALL_IN_RULE_NO)
                            Dim rowSOTSHIPIs() As DataRow = dst.Tables("SOTSHIPI").Select($"CALL_IN_RULE_NO = '{CALL_IN_RULE_NO}'", "", DataViewRowState.OriginalRows)
                            Dim rowSOTSHIPI As DataRow = Nothing
                            If rowSOTSHIPIs.Length > 0 Then
                                rowSOTSHIPI = rowSOTSHIPIs(0)
                            Else
                                rowSOTSHIPI = dst.Tables("SOTSHIPI").Select($"CALL_IN_RULE_NO = '{CALL_IN_RULE_NO}'")(0)
                            End If
                            If Not (rowSOTSHIPI.RowState = DataRowState.Added) Then
                                rowSOTSHIPI.RejectChanges()
                                Dim rowSOTSHIPI_INACTIVE As DataRow = dst.Tables("SOTSHIPI_INACTIVE").NewRow
                                rowSOTSHIPI_INACTIVE.ItemArray = rowSOTSHIPI.ItemArray
                                dst.Tables("SOTSHIPI_INACTIVE").Rows.Add(rowSOTSHIPI_INACTIVE)
                                rowSOTSHIPI_INACTIVE.AcceptChanges()
                                rowSOTSHIPI_INACTIVE.Item("CALL_IN_RULE_STATUS") = "I"
                                rowSOTSHIPI_INACTIVE.Item("LAST_DATE") = DATETIME_STAMP
                                rowSOTSHIPI_INACTIVE.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            End If
                            rowSOTSHIPI.Delete()
                            If rowSOTSHIPIs.Length > 0 Then
                                rowSOTSHIPI.AcceptChanges()
                            End If
                        Next
                        Sort_grdColumns(grdSOTSHIPI_INACTIVE, "SHIP_TO_CODE")
                    End If
                End If

            Case "Re-Activate"

                If grdSOTSHIPI_INACTIVE.Selected.Rows.Count = 0 And grdSOTSHIPI_INACTIVE.ActiveRow IsNot Nothing Then
                    grdSOTSHIPI_INACTIVE.ActiveRow.Selected = True
                End If

                If grdSOTSHIPI_INACTIVE.Selected.Rows.Count > 0 Then

                    Dim CALL_IN_RULE_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPI_INACTIVE.Selected.Rows
                        Dim CALL_IN_RULE_NO As String = grow.Cells("CALL_IN_RULE_NO").Value
                        Dim SHIP_TO_CODE As String = grow.Cells("SHIP_TO_CODE").Value
                        If dst.Tables("SOTSHIPI").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE}'").Length > 0 Then
                            MsgBox($"Cannot Re-Activate Call-In Rule {CALL_IN_RULE_NO}" & vbCrLf & vbCrLf & $"An Active Call-In Rule already exists" & $" for Ship-To {SHIP_TO_CODE}")
                            Exit Sub
                        End If
                        CALL_IN_RULE_NOs.Add(CALL_IN_RULE_NO)
                    Next

                    If MsgBox($"OK to Re-Activate the {CALL_IN_RULE_NOs.Count} selected row(s)?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        For Each CALL_IN_RULE_NO As String In CALL_IN_RULE_NOs
                            Dim rowSOTSHIPI_INACTIVE As DataRow = dst.Tables("SOTSHIPI_INACTIVE").Rows.Find(CALL_IN_RULE_NO)
                            If Not rowSOTSHIPI_INACTIVE.RowState = DataRowState.Added Then
                                Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").NewRow
                                rowSOTSHIPI.ItemArray = rowSOTSHIPI_INACTIVE.ItemArray
                                dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI)
                                rowSOTSHIPI.AcceptChanges()
                                rowSOTSHIPI.Item("CALL_IN_RULE_STATUS") = "A"
                                rowSOTSHIPI.Item("LAST_DATE") = DATETIME_STAMP
                                rowSOTSHIPI.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            End If
                            rowSOTSHIPI_INACTIVE.Delete()
                        Next
                        Sort_grdColumns(grdSOTSHIPI, "SHIP_TO_CODE")
                    End If
                End If

            Case "Create Default Rule for New Ship-Tos", "Copy Rule to New Ship-Tos"

                Dim rowSOTSHIPI As DataRow = Nothing
                If e.Tool.Key = "Copy Rule to New Ship-Tos" Then
                    If grdSOTSHIPI.Selected.Rows.Count = 0 And grdSOTSHIPI.ActiveRow IsNot Nothing Then
                        grdSOTSHIPI.ActiveRow.Selected = True
                    End If

                    If grdSOTSHIPI.Selected.Rows.Count = 0 Then Exit Sub

                    Dim CALL_IN_RULE_NO As String = grdSOTSHIPI.ActiveRow.Cells("CALL_IN_RULE_NO").Value
                    rowSOTSHIPI = dst.Tables("SOTSHIPI").Rows.Find(CALL_IN_RULE_NO)

                End If

                Setup_ARTCUST2_SHIPTO(rowSOTSHIPI)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_TO_KEY"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Validate_SHIP_TO_KEY Then
                        Click_Command("Load")
                    End If
                End If

            Case "SHIP_TO_CODE_PFX"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Validate_SHIP_TO_CODE_PFX() Then
                        Click_Command("Load")
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SHIP_TO_KEY"
                If Validate_SHIP_TO_KEY Then
                    Click_Command("Load")
                End If

            Case "SHIP_TO_CODE_PFX"
                If Validate_SHIP_TO_CODE_PFX() Then
                    Click_Command("Load")
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SHIP_TO_KEY"
                If Validate_SHIP_TO_KEY Then
                    Click_Command("Load")
                End If

            Case "SHIP_TO_CODE_PFX"
                If Validate_SHIP_TO_CODE_PFX() Then
                    Click_Command("Load")
                End If
        End Select
    End Sub

#End Region

    Sub Refresh_Records()
        Fill_Records("SOTSHIPI", New String() {"A", "*"})
        Sort_grdColumns(grdSOTSHIPI, "SHIP_TO_CODE")

        Fix_CheckBoxes()
    End Sub

    Private Sub grdSOTSHIPI_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTSHIPI.InitializeLayout

    End Sub

    Private Sub grdSOTSHIPI_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTSHIPI.DoubleClickRow
        If Not ScreenMode Then
            If e.Row.IsDataRow Then
                Dim SHIP_TO_CODE As String = e.Row.Cells("SHIP_TO_CODE").Value
                Dim rowSOTSHIPD As DataRow = LookUp("SOTSHIPD", SHIP_TO_CODE)
                Dim SHIP_TO_CODE_PFX As String = rowSOTSHIPD.Item("SHIP_TO_CODE_PFX")


                Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)
                Dim SHIP_TO_KEY As String = rowSOTSHIPE.Item("SHIP_TO_KEY")
                Dim SHIP_TO_TYPE As String = rowSOTSHIPE.Item("SHIP_TO_TYPE")

                Absx1.txtFor("SHIP_TO_KEY").Text = SHIP_TO_KEY
                Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX
                Absx1.optFor("SHIP_TO_TYPE").Value = SHIP_TO_TYPE
                ' optSHIP_TO_TYPE.Value = SHIP_TO_TYPE
                Click_Command("Load")
            End If
        End If

    End Sub

    Function Validate_SHIP_TO_KEY() As Boolean

        Dim SHIP_TO_KEY As String = Absx1.txtFor("SHIP_TO_KEY").Text
        Dim SHIP_TO_TYPE As String = Absx1.optFor("SHIP_TO_TYPE").Value

        ASCMAIN1.sql = "Select * from SOTSHIPE where SHIP_TO_TYPE = :PARM1 and SHIP_TO_KEY = :PARM2"
        Dim rowSOTSHIPE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "VV", New String() {SHIP_TO_TYPE, SHIP_TO_KEY})

        If rowSOTSHIPE Is Nothing Then
            Return False
        Else
            Dim SHIP_TO_CODE_PFX As String = rowSOTSHIPE.Item("SHIP_TO_CODE_PFX")
            Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX
            Return True
        End If
    End Function

    Function Validate_SHIP_TO_CODE_PFX() As Boolean

        Dim SHIP_TO_CODE_PFX As String = Absx1.txtFor("SHIP_TO_CODE_PFX").Text

        ASCMAIN1.sql = "Select * from SOTSHIPE where SHIP_TO_CODE_PFX = :PARM1"
        Dim rowSOTSHIPE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {SHIP_TO_CODE_PFX})

        If rowSOTSHIPE Is Nothing Then
            Return False
        Else
            Dim SHIP_TO_KEY As String = rowSOTSHIPE.Item("SHIP_TO_KEY")
            Absx1.txtFor("SHIP_TO_KEY").Text = SHIP_TO_KEY
            Return True
        End If
    End Function

    Sub Setup_ARTCUST2_SHIPTO(rowSOTSHIPI_RULE As DataRow)
        Dim CUST_CODE As String = Absx1.txtFor("SHIP_TO_KEY").Text
        Fill_Records("ARTCUST2_SHIPTO", CUST_CODE)

        For Each T As String In New String() {"SOTSHIPI", "SOTSHIPI_INACTIVE"}
            For Each rowT As DataRow In dst.Tables(T).Select("")
                Dim SHIP_TO_CODE As String = rowT.Item("SHIP_TO_CODE")
                Dim CUST_STORE_NO As String = Mid(SHIP_TO_CODE, 9)
                If CUST_STORE_NO <> "" Then
                    Dim rowARTCUST2_SHIPTO As DataRow = dst.Tables("ARTCUST2_SHIPTO").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                    rowARTCUST2_SHIPTO.Delete()
                End If
            Next
        Next
        dst.Tables("ARTCUST2_SHIPTO").AcceptChanges()

        Dim CUST_STORE_NOs As List(Of String) = Select_ARTCUST2_SHIPTOs(dst.Tables("ARTCUST2_SHIPTO"))

        For Each CUST_STORE_NO As String In CUST_STORE_NOs
            Dim rowARTCUST2_SHIPTO As DataRow = dst.Tables("ARTCUST2_SHIPTO").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            Dim SHIP_TO_CODE As String = rowARTCUST2_SHIPTO.Item("SHIP_TO_CODE")
            Dim CUST_STORE_NAME As String = rowARTCUST2_SHIPTO.Item("CUST_STORE_NAME")
            If rowSOTSHIPI_RULE Is Nothing Then
                Create_Default_CALL_IN_rule(SHIP_TO_CODE, CUST_STORE_NAME)
            Else
                Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").NewRow
                rowSOTSHIPI.ItemArray = rowSOTSHIPI_RULE.ItemArray
                rowSOTSHIPI.Item("CALL_IN_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPI.CALL_IN_RULE_NO")
                rowSOTSHIPI.Item("SHIP_TO_CODE") = SHIP_TO_CODE_PFX & "D" & CUST_STORE_NO
                rowSOTSHIPI.Item("SHIP_TO_NAME") = CUST_STORE_NAME

                rowSOTSHIPI.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPI.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPI.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTSHIPI.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIPI.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPI.Item("LAST_OPER") = ASCMAIN1.USER_ID

                dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI)
            End If
        Next
        Sort_grdColumns(grdSOTSHIPI, "SHIP_TO_CODE")

    End Sub

    Function Select_ARTCUST2_SHIPTOs(tbl As DataTable) As List(Of String)

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO_CALL_IN")
        Dim CUST_STORE_NOs As New List(Of String)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                For Each CUST_STORE_NO As String In ASCMAIN1.CodeSelector.SelectedCodes
                    CUST_STORE_NOs.Add(CUST_STORE_NO)
                Next
            End If
        End If

        Return CUST_STORE_NOs
    End Function

    Sub Create_Default_CALL_IN_rule(SHIP_TO_CODE As String, SHIP_TO_NAME As String)
        Dim rowSOTSHIPI As DataRow = dst.Tables("SOTSHIPI").NewRow
        With rowSOTSHIPI
            .Item("CALL_IN_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPI.CALL_IN_RULE_NO")
            .Item("CALL_IN_RULE_STATUS") = "A"
            .Item("SHIP_TO_CODE") = SHIP_TO_CODE
            .Item("SHIP_TO_NAME") = SHIP_TO_NAME

            .Item("CALL_IN_NOTICE_DAYS") = 3
            .Item("CALL_IN_DAYS") = "0000000"
            .Item("PICK_UP_DAYS") = "0000000"

            .Item("CALL_IN_HOL_EXC") = "E"
            .Item("PICK_UP_HOL_EXC") = "L"

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("SOTSHIPI").Rows.Add(rowSOTSHIPI)
    End Sub

    Function Validate_Active_Records() As String

        Dim EMsg As String = ""

        For Each rowSOTSHIPI As DataRow In dst.Tables("SOTSHIPI").Select

            Dim CALL_IN_RULE_NO As String = rowSOTSHIPI.Item("CALL_IN_RULE_NO") & ""

            Dim CALL_IN_NOTICE_DAYS As String = Val(rowSOTSHIPI.Item("CALL_IN_NOTICE_DAYS") & "")
            If CALL_IN_NOTICE_DAYS < 0 Or CALL_IN_NOTICE_DAYS > 9 Then
                EMsg &= vbCr & $"{CALL_IN_RULE_NO} Call-In Notice must be between 0 and 9 days"
            End If

            Dim CALL_IN_DAYS As String = "0000000"
            Dim PICK_UP_DAYS As String = "0000000"
            With rowSOTSHIPI
                Dim DI As Integer = 0

                For Each D As String In New String() {"SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"}
                    DI += 1
                    If .Item($"CALL_IN_{D}") = "1" Then MID(CALL_IN_DAYS, DI, 1) = "1"
                    If .Item($"PICK_UP_{D}") = "1" Then MID(PICK_UP_DAYS, DI, 1) = "1"
                Next
            End With

            Mid(CALL_IN_DAYS, 1, 1) = "0"
            Mid(CALL_IN_DAYS, 7, 1) = "0"
            Mid(PICK_UP_DAYS, 1, 1) = "0"
            Mid(PICK_UP_DAYS, 7, 1) = "0"

            rowSOTSHIPI.Item("CALL_IN_DAYS") = CALL_IN_DAYS
            rowSOTSHIPI.Item("PICK_UP_DAYS") = PICK_UP_DAYS
        Next

        Return EMsg
    End Function

    Private Sub grdSOTSHIPI_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTSHIPI.BeforeRowUpdate

    End Sub

    Private Sub grdSOTSHIPI_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdSOTSHIPI.BeforeCellUpdate

    End Sub

    Private Sub grdSOTSHIPI_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs) Handles grdSOTSHIPI.BeforeExitEditMode

    End Sub

    Sub Fix_CheckBoxes()

        For Each T As String In New String() {"SOTSHIPI", "SOTSHIPI_INACTIVE"}
            For Each row As DataRow In dst.Tables(T).Select("")
                Dim CALL_IN_DAYS As String = row.Item("CALL_IN_DAYS") & ""
                If CALL_IN_DAYS = "" Then CALL_IN_DAYS = "0000000"
                Dim PICK_UP_DAYS As String = row.Item("PICK_UP_DAYS") & ""
                If PICK_UP_DAYS = "" Then PICK_UP_DAYS = "0000000"

                Dim DI As Integer = 0
                For Each D As String In New String() {"SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"}
                    DI += 1
                    row.Item($"CALL_IN_{D}") = (Mid(CALL_IN_DAYS, DI, 1))
                    row.Item($"PICK_UP_{D}") = (Mid(PICK_UP_DAYS, DI, 1))
                Next
            Next
        Next

    End Sub
End Class