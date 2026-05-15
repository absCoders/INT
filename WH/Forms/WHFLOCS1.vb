Public Class WHFLOCS1

    Dim sqlWHTLOCB2where As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select X.WHSE_CODE, X.LOCATION_CODE, X.TRANS, X.INIT_DATE" _
                & " from WHTLOCM1,(" _
                & "Select WHSE_CODE, LOCATION_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 group by WHSE_CODE, LOCATION_CODE" _
                & ") X where WHTLOCM1.WHSE_CODE = X.WHSE_CODE and WHTLOCM1.LOCATION_CODE = X.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCMA", "**", 0, False, "", 2)
            .Tables("WHTLOCMA").Columns("TRANS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM1"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select X.ITEM_CODE, ICTITEM1.ITEM_DESC, X.TRANS, X.INIT_DATE" _
                & " from ICTITEM1,(" _
                & "Select ITEM_CODE, COUNT (*) TRANS, MAX (INIT_DATE) INIT_DATE" _
                & " from WHTLOCB2 where INIT_DATE > SYSDATE - 24 group by ITEM_CODE" _
                & ") X where ICTITEM1.ITEM_CODE = X.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTYLA", "**", 0, False, "", 1)
            .Tables("ICTSTYLA").Columns("TRANS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "", 4)
            .Tables("WHTLOCB1").Columns.Add("PERSIST")

            ASCMAIN1.sql = "Select WHTLOCB2.* from WHTLOCB2"
            Create_TDA(.Tables.Add, "WHTLOCB2", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select WHTMOVE2.*" & vbCrLf _
                & ", WHTMOVE1.WHSE_TRAN_TYPE, WHTMOVE1.WHSE_CODE, WHTMOVE1.SESSION_NO" & vbCrLf _
                & " from WHTMOVE1, WHTMOVE2" & vbCrLf _
                & " where WHTMOVE1.WHSE_TRAN_NO = WHTMOVE2.WHSE_TRAN_NO" & vbCrLf _
                & "   and WHTMOVE2.WHSE_TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTMOVEX", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1'"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
        End With

        Fill_Records("ICTWHSE1")

        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")
        grdWHTLOCB2.DataSource = dst.Tables("WHTLOCB2")
        grdWHTMOVEX.DataSource = dst.Tables("WHTMOVEX")
        grdWHTLOCM1.DataSource = dst.Tables("WHTLOCM1")
        grdWHTLOCMA.DataSource = dst.Tables("WHTLOCMA")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")
        grdICTITEMA.DataSource = dst.Tables("ICTSTYLA")

        Create_Summary(grdWHTLOCB1, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB1, "ITEM_CODE", "Count")
        Create_Summary(grdWHTLOCB1, "LOCATION_QTY")

        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_TYPE", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_QTY")

        Create_Summary(grdWHTMOVEX, "WHSE_TRAN_LNO", "Count")
        Create_Summary(grdWHTMOVEX, "WHSE_TRAN_QTY")

        With grdWHTLOCB1.DisplayLayout.Bands(0)
            .Columns("LOCATION_CODE").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("ITEM_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("LAST_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("LAST_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        With grdWHTLOCB2.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("WHSE_TRAN_LNO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_CODE_OTHER").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("WHSE_TRAN_TYPE").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("WHSE_TRAN_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("INIT_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("INIT_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        With grdWHTMOVEX.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("WHSE_TRAN_LNO").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("LOCATION_CODE_FROM").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("LOCATION_CODE_TO").Header.Appearance.BackColor2 = Drawing.Color.Gold
            .Columns("WHSE_TRAN_QTY").Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            .Columns("ITEM_CODE").Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            .Columns("INIT_OPER").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("INIT_DATE").Header.Appearance.BackColor2 = Drawing.Color.Orange
        End With

        Dim WH_PARM_BARCODE_SUPPORT As String = "0" ' PUT THIS IN WHTPARM1 AS A CHECKBOX
        grdWHTLOCB1.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = (WH_PARM_BARCODE_SUPPORT <> "1")
        grdWHTLOCB2.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = (WH_PARM_BARCODE_SUPPORT <> "1")
        grdWHTMOVEX.DisplayLayout.Bands(0).Columns("BAR_CODE").Hidden = (WH_PARM_BARCODE_SUPPORT <> "1")

        ASCMAIN1.Add_Value_List(grdWHTLOCB2, "WHSE_TRAN_TYPE", Nothing, New String() {":", "R:Receipt", "P:PhysAdj", "D:?", "M:Move", "C:Cartonize", "U:Usage", "A:Adj", "S:SlsShp", "L:?", "T:WhsXfr", "N:Returns", "X:XfrRec"})
        ' PULL TO, MOVE FROM

        grpStyle.Top = grpLocation.Top
        grpStyle.Left = grpLocation.Left

        ' Integrity_Check(False) ' NO WHSE YET
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                Validate_Code("WHSE_CODE")
                If cdr IsNot Nothing Then
                    If cdr.Item("WHSE_LOCATOR") & "" <> "1" Then
                        EMsg &= vbCr & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " is NOT set up for Location"
                    End If
                End If

                If optViewBy.Value <> "S" Then
                    If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must First Specify a Location"
                    Else
                        Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() _
                                                            {Absx1.txtFor("WHSE_CODE").Text, _
                                                             Absx1.txtFor("LOCATION_CODE").Text})
                        If rowWHTLOCM1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Location (" & Absx1.txtFor("LOCATION_CODE").Text & ")"
                        End If
                    End If
                End If

                If optViewBy.Value = "S" Then
                    If Absx1.txtFor("ITEM_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must First Specify a Item"
                    Else
                        Absx1.txtFor("ITEM_CODE").Text = Absx1.txtFor("ITEM_CODE").Text.ToUpper
                        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", New String() _
                                                            {Absx1.txtFor("ITEM_CODE").Text})
                        If rowICTITEM1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Item (" & Absx1.txtFor("ITEM_CODE").Text & ")"
                        End If
                    End If
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

            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Print"
                'Print_Record()

            Case "Done"
                Mode_Settings(False)

            Case "Integrity Check"
                Integrity_Check(True)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Select").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Integrity Check").Settings.Enabled = not_iScreenMode

                    .Items("Print").Visible = False

                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splWHTLOCBX.Visible = tf
        splVisited.Visible = Not tf

        If ScreenMode Then
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("LOCATION_CODE").Hidden = Not (optViewBy.Value = "S")
            grdWHTLOCB1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Hidden = (optViewBy.Value = "S")

        Else
            Clear_Record()
            Setup_ViewBy()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"WHTLOCB1", "WHTLOCB2", "WHTMOVEX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Absx1.txtFor("ITEM_CODE").Clear()
        Absx1.txtFor("LOCATION_CODE").Clear()
        Absx1.txtFor("LOCATION_CODE2").Clear()

        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If dst.Tables("ICTWHSE1").Rows.Count = 1 Then
                Absx1.txtFor("WHSE_CODE").Text = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_CODE")
            End If
        End If

        Fill_Records("ICTSTYLA")
        Fill_Records("WHTLOCMA")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        sqlWHTLOCB2where = ""

        If optViewBy.Value = "S" Then
            Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If dst.Tables("ICTITEM1").Rows.Find(New String() {ITEM_CODE}) Is Nothing Then
                dst.Tables("ICTITEM1").Rows.Add(New String() {ITEM_CODE, rowICTITEM1.Item("ITEM_DESC") & ""})
            End If
            Sort_grdColumns(grdICTITEM1, "ITEM_CODE")

            sqlWHTLOCB2where = "" _
                & " where ITEM_CODE = '" & ITEM_CODE & "'" _
                & " and WHSE_CODE = '" & WHSE_CODE & "'"

            grdWHTLOCB1.Text = "Locations containing Style " & ITEM_CODE

        Else
            Dim LOCATION_CODE As String = Absx1.txtFor("LOCATION_CODE").Text
            If dst.Tables("WHTLOCM1").Rows.Find(New String() {WHSE_CODE, LOCATION_CODE}) Is Nothing Then
                dst.Tables("WHTLOCM1").Rows.Add(New String() {WHSE_CODE, LOCATION_CODE})
            End If
            Sort_grdColumns(grdWHTLOCM1, "WHSE_CODE, LOCATION_CODE")

            sqlWHTLOCB2where = "" _
                & " where LOCATION_CODE = '" & LOCATION_CODE & "'" _
                & " and WHSE_CODE = '" & WHSE_CODE & "'"

            grdWHTLOCB1.Text = "Styles in Location " & LOCATION_CODE
        End If

        ASCMAIN1.sql = "Select * from WHTLOCB1" & sqlWHTLOCB2where
        Fill_Records("WHTLOCB1", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where
        Fill_Records("WHTLOCB2", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdWHTLOCB1, "LOCATION_CODE,ITEM_CODE")

        Setup_grdWHTLOCB1()
        Setup_grdWHTLOCB2()
        Setup_grdWHTMOVEX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where = "WHSE_LOCATOR = '1'"
        End Select

    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Select"
                If key.StartsWith("S:") Then
                    optViewBy.Value = "S"
                    Absx1.txtFor("ITEM_CODE").Text = Split(key, ":")(1)
                    Click_Command("Select")
                ElseIf key.StartsWith("L:") Then
                    optViewBy.Value = "L"
                    Absx1.txtFor("LOCATION_CODE").Text = Split(key, ":")(1)
                    Click_Command("Select")
                End If

        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTLOCB1, "SSSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show 0 Qty", "Move to ...", "Move to Bin", "Location Inquiry")
        Load_Popup_Menu(grdWHTLOCB2, "BB", "Reverse Entire Move (All Lines shown Below)", "Reverse This Move Line Only")
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
            'e.Cancel = True
        Else
            Dim allow_reverse As Boolean = (ASCMAIN1.USER_SECURITY_CODEs.Contains("WH"))
            If grd.Name = "grdWHTLOCB2" Then
                allow_reverse = allow_reverse AndAlso grd.ActiveRow IsNot Nothing And grd.ActiveRow.Cells("WHSE_TRAN_TYPE").Value = "M"
            End If

            Select Case e.SourceControl.Name
                Case "grdWHTLOCB1"
                    tlb_pop.Tools("Move to ...").SharedProps.Visible = allow_reverse
                Case "grdWHTLOCB2"
                    tlb_pop.Tools("Reverse Entire Move (All Lines shown Below)").SharedProps.Visible = allow_reverse
                    tlb_pop.Tools("Reverse This Move Line Only").SharedProps.Visible = allow_reverse
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Show 0 Qty"
                Setup_grdWHTLOCB1()

            Case "Move to ...", "Move to Bin"
                If grdWHTLOCB1.Selected.Rows.Count = 0 Then
                    grdWHTLOCB1.ActiveRow.Selected = True
                Else
                    If Not grdWHTLOCB1.ActiveRow.Selected Then
                        Exit Sub
                    End If
                End If
                Move_To(e.Tool.Key = "Move to Bin")


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdWHTLOCB1" Then

            Select Case e.Tool.Key
                Case "Location Inquiry"
                    Dim KEY As String = ""
                    If optViewBy.Value = "L" Then
                        Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                        KEY = "S:" & ITEM_CODE
                    Else
                        Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Value
                        KEY = "L:" & LOCATION_CODE
                    End If

                    Context_Launch("Select", KEY, e.Tool.Key, "WHFLOCS1")
            End Select
        End If

        If grd.Name = "grdWHTLOCB2" Then

            Select Case e.Tool.Key
                Case "Reverse Entire Move (All Lines shown Below)"
                    Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
                    Reverse_Transaction(WHSE_TRAN_NO)

                Case "Reverse This Move Line Only"
                    Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
                    Dim WHSE_TRAN_LNO As Int32 = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_LNO").Value & ""
                    Reverse_Transaction(WHSE_TRAN_NO, WHSE_TRAN_LNO)
            End Select
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LOCATION_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select", e)
                End If
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "LOCATION_CODE"
                Click_Command("Select")
            Case "ITEM_CODE"
                Click_Command("Select")
        End Select
    End Sub

#End Region

#Region "grdWHTLOCB1"
    Private Sub grdWHTLOCB1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTLOCB1.AfterRowActivate
        Setup_grdWHTLOCB2()
    End Sub
#End Region

    Sub Setup_grdWHTLOCB1()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show 0 Qty"), UltraWinToolbars.StateButtonTool)
        Dim dvw As DataView = DirectCast(grdWHTLOCB1.DataSource, DataTable).DefaultView
        If Not tlb_sbt.Checked Then
            dvw.RowFilter = "LOCATION_QTY <> 0 OR PERSIST = '1'"
        Else
            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("PERSIST = '1'")
                rowWHTLOCB1.Item("PERSIST") = "0"
            Next
            dvw.RowFilter = ""
        End If
    End Sub

    Sub Setup_grdWHTLOCB2()
        If grdWHTLOCB1.ActiveRow Is Nothing Then
            grdWHTLOCB2.Visible = False
        Else
            grdWHTLOCB2.Visible = True
            Dim LOCATION_CODE As String = grdWHTLOCB1.ActiveRow.Cells("LOCATION_CODE").Value
            Dim ITEM_CODE As String = grdWHTLOCB1.ActiveRow.Cells("ITEM_CODE").Value
            Dim DVW As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
            DVW.RowFilter = "LOCATION_CODE = '" & LOCATION_CODE & "' and ITEM_CODE = '" & ITEM_CODE & "'"
            grdWHTLOCB2.Text = "Audit Trail for Location " & LOCATION_CODE & ", Style " & ITEM_CODE
            Sort_grdColumns(grdWHTLOCB2, "INIT_DATE")
        End If
    End Sub

    Sub Setup_grdWHTMOVEX()
        If grdWHTLOCB2.ActiveRow Is Nothing Then
            grdWHTMOVEX.Visible = False
        Else
            grdWHTMOVEX.Visible = True
            Dim WHSE_TRAN_NO As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""
            Dim WHSE_TRAN_TYPE As String = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_TYPE").Value & ""
            If WHSE_TRAN_TYPE <> "M" Then
                grdWHTMOVEX.Visible = False
            Else
                Fill_Records("WHTMOVEX", WHSE_TRAN_NO)
                Sort_grdColumns(grdWHTMOVEX, "WHSE_TRAN_LNO")
                grdWHTMOVEX.Text = "Transaction Details for " & WHSE_TRAN_NO
            End If
        End If
    End Sub

    Sub Integrity_Check(ack_if_ok As Boolean)

        If Absx1.txtFor("WHSE_CODE").Text = "" Then Exit Sub

        ASCMAIN1.sql = "Select ITEM_CODE, LOC, STY, VAR, LOC - (STY + VAR) OUT FROM (" & vbCrLf _
            & " Select ITEM_CODE, SUM (LOC) LOC, SUM (STY) STY, SUM (VAR) VAR FROM (" & vbCrLf _
            & " Select ITEM_CODE, SUM (LOCATION_QTY) LOC, 0 STY, 0 VAR" & vbCrLf _
            & " from WHTLOCB1" & vbCrLf _
            & " where WHTLOCB1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by ITEM_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select ITEM_CODE, 0 LOC, SUM (WHSE_QTY_ON_HAND) STY, 0 VAR" & vbCrLf _
            & " from ICTSTAT2 where ICTSTAT2.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'" & vbCrLf _
            & " group by ITEM_CODE" & vbCrLf _
            & " ) group by ITEM_CODE" & vbCrLf _
            & " ) where LOC - (STY + VAR) <> 0"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count = 0 Then
            If ack_if_ok Then
                MsgBox("All OK", MsgBoxStyle.OkOnly, "Verification")
            End If
        Else
            Using F As New ASFMSGBF
                F.Show_grd(tbl, Me, "Items Out of Balance (Locator vs Perpetual)")
            End Using
        End If
    End Sub

    Sub Setup_ViewBy()
        grpStyle.Visible = (optViewBy.Value = "S")
        grpLocation.Visible = Not (optViewBy.Value = "S")
        lblLOCATION_CODE2.Visible = (optViewBy.Value = "R")
        txtLOCATION_CODE2.Visible = (optViewBy.Value = "R")
    End Sub

    Private Sub optViewBy_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optViewBy.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ViewBy()
    End Sub

    Private Sub grdICTITEM1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEM1.DoubleClickRow
        optViewBy.Value = "S"
        'Absx1.txtFor("WHSE_CODE").Text = grdICTITEM1.ActiveRow.Cells("WHSE_CODE").Text
        Absx1.txtFor("ITEM_CODE").Text = grdICTITEM1.ActiveRow.Cells("ITEM_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCM1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCM1.DoubleClickRow
        If grdWHTLOCM1.ActiveRow IsNot Nothing AndAlso grdWHTLOCM1.ActiveRow.IsDataRow Then
            optViewBy.Value = "L"
            Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCM1.ActiveRow.Cells("WHSE_CODE").Text
            Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCM1.ActiveRow.Cells("LOCATION_CODE").Text
            Click_Command("Select")
        End If
    End Sub

    Private Sub grdICTSTYLA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEMA.DoubleClickRow
        optViewBy.Value = "S"
        'Absx1.txtFor("WHSE_CODE").Text = grdICTSTYLA.ActiveRow.Cells("WHSE_CODE").Text
        Absx1.txtFor("ITEM_CODE").Text = grdICTITEMA.ActiveRow.Cells("ITEM_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCMA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTLOCMA.DoubleClickRow
        optViewBy.Value = "L"
        Absx1.txtFor("WHSE_CODE").Text = grdWHTLOCMA.ActiveRow.Cells("WHSE_CODE").Text
        Absx1.txtFor("LOCATION_CODE").Text = grdWHTLOCMA.ActiveRow.Cells("LOCATION_CODE").Text
        Click_Command("Select")
    End Sub

    Private Sub grdWHTLOCB2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLOCB2.AfterRowActivate
        Setup_grdWHTMOVEX()
    End Sub

    Private Sub grdWHTPULLX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTMOVEX.InitializeLayout

    End Sub

    Private Sub grdWHTPULLX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTMOVEX.InitializeRow
        If e.Row.Cells("WHSE_TRAN_LNO").Value = grdWHTLOCB2.ActiveRow.Cells("WHSE_TRAN_LNO").Value Then
            e.Row.Appearance.ForeColor = Drawing.Color.Blue
        End If
        If e.Row.Cells("STATUS").Value & "" = "R" Then
            e.Row.Cells("WHSE_TRAN_LNO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("WHSE_TRAN_LNO").ToolTipText = "This Line was Reversed"
        End If
    End Sub

    Sub Reverse_Transaction(WHSE_TRAN_NO As String, Optional WHSE_TRAN_LNO As Int32 = 0)
        If Not ASCMAIN1.Logical_Lock("WHTMOVE1", WHSE_TRAN_NO, , , , 1) Then Exit Sub

        Dim rowWHTMOVE1 As DataRow = Fill_Record("WHTMOVE1", WHSE_TRAN_NO)
        If rowWHTMOVE1.Item("STATUS") & "" = "R" Then
            MsgBox("Move Transaction " & WHSE_TRAN_NO & " has already been reversed", _
                   MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If MsgBox("Are you sure that you want to Reverse Move Transaction " _
                  & WHSE_TRAN_NO & IIf(WHSE_TRAN_LNO = 0, "", ", Line " & CStr(WHSE_TRAN_LNO)), _
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        BeginTrans()
        rowWHTMOVE1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        Update_Record_TDA("WHTMOVE1")
        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, WHSE_TRAN_LNO, -1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
        CommitTrans()

        Dim sqlw As String = " and WHSE_TRAN_TYPE = 'M' and WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'" _
                             & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO))
        For Each rowWHTLOCB2 As DataRow In dst.Tables("WHTLOCB2").Select(Mid(sqlw, 5))
            rowWHTLOCB2.Delete()
        Next

        ASCMAIN1.sql = "Select * from WHTLOCB2" & sqlWHTLOCB2where & sqlw _
                             & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO))

        Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from WHTLOCB1" _
            & sqlWHTLOCB2where _
            & " and (LOCATION_CODE,ITEM_CODE) in" _
            & " (Select LOCATION_CODE,ITEM_CODE from WHTLOCB2 " _
            & sqlWHTLOCB2where & sqlw _
            & IIf(WHSE_TRAN_LNO = 0, "", " and WHSE_TRAN_LNO = " & CStr(WHSE_TRAN_LNO)) & ")"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select
            Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New String() { _
                                                        row.Item("WHSE_CODE"), _
                                                        row.Item("LOCATION_CODE"), _
                                                        row.Item("BAR_CODE"), _
                                                        row.Item("ITEM_CODE")})
            rowWHTLOCB1.Item("LOCATION_QTY") = row.Item("LOCATION_QTY")
            rowWHTLOCB1.Item("LAST_DATE") = row.Item("LAST_DATE")
            rowWHTLOCB1.Item("LAST_OPER") = row.Item("LAST_OPER")
            rowWHTLOCB1.Item("PERSIST") = "1"
        Next

        ' Setup_grdWHTLOCB1()
        Setup_grdWHTLOCB2()
        Setup_grdWHTMOVEX()

        ASCMAIN1.MultiTask_Release(, , 1)

        MsgBox("Move Transaction " & WHSE_TRAN_NO _
               & IIf(WHSE_TRAN_LNO = 0, "", ", Line " & CStr(WHSE_TRAN_LNO)) _
               & " has been Successfully Reversed", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Move_To(Optional move_to_bin As Boolean = False)

        Using ff As New TAC.TAFLOCM1()
            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTLOCB1.Selected.Rows
                Dim LOCATION_CODE_TO As String = ""
                If move_to_bin Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", row.Cells("ITEM_CODE").Value)
                    If rowICTITEM1.Item("ITEM_BIN") & "" <> "" AndAlso _
                        LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, rowICTITEM1.Item("ITEM_BIN") & ""}) IsNot Nothing Then
                        LOCATION_CODE_TO = rowICTITEM1.Item("ITEM_BIN") & ""
                    End If
                End If

                'row.Cells("LOCATION_QTY").Value

                ff.AddItemToMove(row.Cells("WHSE_CODE").Value, _
                                 row.Cells("LOCATION_CODE").Value, _
                                 row.Cells("ITEM_CODE").Value, _
                                 0, _
                                 LOCATION_CODE_TO)
            Next

            ff.ShowDialog()

            Dim WHSE_TRAN_NO As String = ff.WHSE_TRAN_NO
            If WHSE_TRAN_NO <> "" Then

                ASCMAIN1.sql = "Select * from WHTLOCB2" _
                    & sqlWHTLOCB2where _
                    & " and WHSE_TRAN_TYPE = 'M' and WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"

                Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)

                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWHTLOCB1.Selected.Rows

                    Dim rowWHTLOCB1 As DataRow = LookUp("WHTLOCB1", New String() { _
                                                        row.Cells("WHSE_CODE").Value, _
                                                        row.Cells("LOCATION_CODE").Value, _
                                                        row.Cells("BAR_CODE").Value, _
                                                        row.Cells("ITEM_CODE").Value})
                    row.Cells("LOCATION_QTY").Value = rowWHTLOCB1.Item("LOCATION_QTY")
                    row.Cells("LAST_DATE").Value = rowWHTLOCB1.Item("LAST_DATE")
                    row.Cells("LAST_OPER").Value = rowWHTLOCB1.Item("LAST_OPER")
                    row.Cells("PERSIST").Value = "1"
                    row.Update()
                Next
            End If

            Dim SAVE As String = ""
            If optViewBy.Value = "S" Then
                SAVE = Absx1.txtFor("ITEM_CODE").Text
            Else
                SAVE = Absx1.txtFor("LOCATION_CODE").Text
            End If
            Click_Command("Done")
            If optViewBy.Value = "S" Then
                Absx1.txtFor("ITEM_CODE").Text = SAVE
            Else
                Absx1.txtFor("LOCATION_CODE").Text = SAVE
            End If
            Click_Command("Select")

            'setup_grdWHTLOCB2()
            'Setup_grdWHTMOVEX()

        End Using
    End Sub

    Private Sub grdICTSTYL1_DoubleClickRow(sender As System.Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYL1.DoubleClickRow, grdICTITEM1.DoubleClickRow

    End Sub
End Class