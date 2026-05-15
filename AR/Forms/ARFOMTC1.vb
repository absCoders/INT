Public Class ARFOMTC1
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim COLUMN_NAMEs_to_edit As String = "INV_DATE,INV_DUE_DATE,INV_CUST_PO,REASON_CODE,LAST_OPER,LAST_DATE,INV_NOTES"
        If ASCMAIN1.CLIENT = "INT" Then
            COLUMN_NAMEs_to_edit = "INV_CUST_PO,REASON_CODE,LAST_OPER,LAST_DATE,INV_NOTES"
        End If

        Get_PARM("ARTPARM1")
        With dst
            'ASCMAIN1.sql = "Select ARTOPEN1.* from ARTOPEN1 where CUST_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ARTOPEN1", "**", 0, True, "V", 3)
            Create_TDA(.Tables.Add, "ARTOPEN1", "*", 1, True, , , COLUMN_NAMEs_to_edit)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)

            Create_TDA(.Tables.Add, "ARTREAS1", "*", 0, False)

            ASCMAIN1.sql = "Select ASTAUDT1.* from ASTAUDT1 where INIT_DATE > SYSDATE -90 and TABLE_NAME = 'ARTOPEN1'"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False, "", 0)

        End With

        Fill_Record("ARTREAS1")

        grdARTOPEN1.DataSource = dst.Tables("ARTOPEN1")
        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")

        Create_Summary(grdARTOPEN1, "INV_NUM", "Count")
        Create_Summary(grdARTOPEN1, New String() {"INV_BALANCE", "INV_TOTAL_AMOUNT"})

        grdARTOPEN1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdARTOPEN1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdARTOPEN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        With grdARTOPEN1.DisplayLayout.Bands("ARTOPEN1")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Split(COLUMN_NAMEs_to_edit, ",").Contains(gcol.Key) Or New String() {"INV_TYPE", "INV_NUM", "CUST_STORE_NO", "INV_BALANCE", "INV_TOTAL_AMOUNT", "INV_DATE"}.Contains(gcol.Key) Then
                    gcol.Hidden = False
                Else
                    gcol.Hidden = True
                End If

                If New String() {"LAST_OPER", "LAST_DATE"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If


                If Split(COLUMN_NAMEs_to_edit, ",").Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke

                End If
            Next

            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"INV_TOTAL_AMOUNT", "INV_BALANCE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In Split(COLUMN_NAMEs_to_edit, ",")
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTCUST1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                Dim BAD_REASON_CODEs As New List(Of String)
                For Each row As DataRow In dst.Tables("ARTOPEN1").Select("", "", DataViewRowState.ModifiedCurrent)
                    Dim REASON_CODE As String = row.Item("REASON_CODE") & ""
                    Dim INV_TYPE As String = row.Item("INV_TYPE")
                    If INV_TYPE = "I" And REASON_CODE = "SHP" Then
                    Else
                        If REASON_CODE = "" Then
                            If INV_TYPE = "B" Or INV_TYPE = "C" Or INV_TYPE = "D" Or INV_TYPE = "O" Then
                                If Not BAD_REASON_CODEs.Contains(REASON_CODE) Then
                                    EMsg &= vbCr & "Reason Code required for Types B, C, D and O"
                                    BAD_REASON_CODEs.Add(REASON_CODE)
                                End If
                            End If
                        Else
                            Dim rowARTREAS1 As DataRow = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                            If rowARTREAS1 Is Nothing Then
                                If Not BAD_REASON_CODEs.Contains(REASON_CODE) Then
                                    EMsg &= vbCr & "Invalid Reason Code in Open Invoice grid (" & REASON_CODE & ")"
                                    BAD_REASON_CODEs.Add(REASON_CODE)
                                End If
                            End If
                        End If
                    End If
                Next
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
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdARTOPEN1.Visible = ScreenMode
        grdASTAUDTX.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTOPEN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("ASTAUDTX")
        Sort_grdColumns(grdASTAUDTX, "INIT_DATE".ToLower)

        CUST_CODE = ""

        Absx1.txtFor("CUST_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")

        EnforceConstraints(False)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        Fill_Records("ARTOPEN1", CUST_CODE)
        EnforceConstraints(True)

        Setup_grdARTOPEN1()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        WriteAuditTrail("ARTOPEN1")
        Update_Record_TDA("ARTOPEN1")
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        '"Retrieve Paid Invoices"
        '"Show $0 Balance Items"
        Load_Popup_Menu(grdARTOPEN1, "SSBBB", "Show Filter", "Show GroupBox", "Show", "Sales Order Inquiry", "Customer Returns Inquiry")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdARTOPEN1"

                    If grd.ActiveRow IsNot Nothing Then
                        If grd.ActiveRow.Band.Index <> 0 Then
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                 
                    tlb_pop.Tools("Sales Order Inquiry").SharedProps.Visible = (grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("INV_TYPE").Text = "I")
                    tlb_pop.Tools("Customer Returns Inquiry").SharedProps.Visible = (grd.ActiveRow.IsDataRow AndAlso grd.ActiveRow.Cells("INV_TYPE").Text = "R")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show $0 Balance Items"
                ' Filter_ARTOPEN1()
            Case "Retrieve Paid Invoices"
                'Dim numDays As Double = 0
                'Using FRM As New ASFMSGBF
                '    numDays = FRM.Get_numint_from_User("Days to Retrieve", "Retrieve Paid Invoices")
                '    If FRM.user_option <> -1 Then
                '        Retrieve_Paid_Invoices(numDays)
                '    End If
                'End Using

            Case "Retrieve Paid Items"
                'Dim numDays As Int64 = 60
                'Retrieve_Paid_Items(60)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Show"
                If grd.ActiveRow.Cells("INV_TYPE").Value & "" <> "I" Then Exit Sub
                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NUM").Value & ""
                Dim FILENAME As String = Create_Invoice(INV_NO)
                Show_Document(FILENAME)

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = ""
                If e.Tool.OwningMenu.Key = "grdARTOPEN1" Then
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Else
                    ORDR_NO = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                End If
                If ORDR_NO <> "" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")
                End If
                 
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs) Handles tlb.ToolValueChanged

    End Sub

#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub
#End Region

#Region "grdRSTBUDR1"

    Private Sub grdRSTBUDR1_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTOPEN1.AfterExitEditMode
        With grdARTOPEN1
            Select Case .ActiveCell.Column.Key
                Case "REASON_CODE"
                    Dim REASON_CODE As String = .ActiveCell.Text
                    Dim INV_TYPE As String = .ActiveCell.Row.Cells("INV_TYPE").Value & ""
                    If INV_TYPE = "I" And REASON_CODE = "SHP" Then
                    Else
                        If REASON_CODE <> "" Then
                            .ActiveCell.Value = ASCMAIN1.Format_Field(REASON_CODE, .ActiveCell.Column.Key)
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTBUDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTOPEN1.AfterRowActivate
        'With grdARTOPEN1.DisplayLayout.Bands(0)
        '    If grdARTOPEN1.ActiveRow.IsAddRow Then
        '        .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Else
        '        .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdRSTBUDR1_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTOPEN1.BeforeExitEditMode
        'e.Cancel = True
    End Sub

    Private Sub grdRSTBUDR1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTOPEN1.BeforeRowUpdate
        With grdARTOPEN1
            'If Val(e.Row.Cells("ORDR_QTY").Text) < 0 Then
            '    MsgBox("Invalid Value entered for Order Qty (" & e.Row.Cells("ORDR_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If

            Dim REASON_CODE As String = e.Row.Cells("REASON_CODE").Text
            Dim INV_TYPE As String = e.Row.Cells("INV_TYPE").Text
            If INV_TYPE = "I" And REASON_CODE = "SHP" Then
            Else
                If INV_TYPE = "B" Or INV_TYPE = "C" Or INV_TYPE = "D" Or INV_TYPE = "O" Then
                    If dst.Tables("ARTREAS1").Rows.Find(REASON_CODE) Is Nothing Then
                        MsgBox("Invalid Value entered for Reason Code (" & REASON_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            'If e.Cancel Then
            '    e.Row.CancelUpdate()
            'End If

            'If Not e.Cancel Then
            '    If e.Row.Cells("OPS_YYYY").Text = "" And ScreenMode Then
            '        .ActiveRow.Cells("OPS_YYYY").Value = Absx1.CtlFor("OPS_YYYY").Text
            '        .ActiveRow.Cells("CUST_CODE").Value = Absx1.CtlFor("CUST_CODE").Text
            '    End If
            'End If
        End With
    End Sub

    Private Sub grdRSTBUDR1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPEN1.ClickCellButton
        Select Case grdARTOPEN1.ActiveCell.Column.Key
            Case "REASON_CODE"
                grdClickCellButton(grdARTOPEN1)
        End Select
    End Sub

    Private Sub grdRSTBUDR1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTOPEN1.InitializeLayout
        'e.Layout.Override.AllowMultiCellOperations = Infragistics.Win.UltraWinGrid.AllowMultiCellOperation.All
        ' e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
    End Sub

#End Region

    'Private Sub optBP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If SELECTION_NO = 0 Then Exit Sub
    '    Setup_grdARTOPEN1()
    'End Sub

    Sub Setup_grdARTOPEN1()
        grdARTOPEN1.Text = "Open AR Items for " & CUST_CODE
        Sort_grdColumns(grdARTOPEN1, "INV_TYPE,INV_NUM")
    End Sub

    Function Create_Invoice(ByVal INV_NO As String) As String
        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.Progress("Now Preparing Invoice for Printing")

        Dim REPORT_NAME As String = "SORINVP1"
        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then RPT = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim sql As String = " and (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO) in ('I'," & INV_NO & ")"
        Dim tempFileName As String = "INV" & DateTime.Now.ToString("yyyyMMddHHmmss")

        REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
        Dim FILENAME As String = ""
        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", "")
            Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
            FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            .Print_Report_End(, True)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return FILENAME
    End Function
End Class