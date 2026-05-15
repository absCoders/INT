Public Class ICFVARS1

    Dim ICTVARS1 As String = ""

    Dim rowICTCOSTM As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFVARS1" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")

        Create_WorkTables(True)

        With dst
            ASCMAIN1.sql = $"Select * from {ICTVARS1}"
            Create_TDA(.Tables.Add, "ICTVARS1", "**", 0, False)
            With .Tables("ICTVARS1")
                '.Columns.Add("SEL")
                '.Columns("SEL").DefaultValue = "0"
                .Columns.Add("TRAN_PV_ORIG", GetType(System.Decimal))
                .Columns.Add("TRAN_MV_ORIG", GetType(System.Decimal))
                .Columns.Add("TRAN_FV_ORIG", GetType(System.Decimal))
                .Columns.Add("TRAN_TV_ORIG", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "ICTFRTC1", "*", 0, False)
            Fill_Records("ICTFRTC1")
            Create_TDA(.Tables.Add, "ICTTRFC1", "*", 0, False)
            Fill_Records("ICTTRFC1")

            ASCMAIN1.sql = "Select POTORDR1.* from POTORDR1 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, True, "V")
            ASCMAIN1.sql = "Select ICTIREC1.* from ICTIREC1 where RECEIPT_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIREC1", "**", 0, True, "V", , "TRF_CLASS_CODE, FRT_CLASS_CODE, TRF_CLASS_PCT, FRT_CLASS_PCT")

            AUDIT.Add("ICTIREC1", "E")

            ASCMAIN1.sql = "Select ICTCOSTA.* from ICTCOSTA"
            Create_TDA(.Tables.Add, "ICTCOSTM", "**", 0, False, "", 0)
            .Tables("ICTCOSTM").Columns.Add("SEQ")
            .Tables("ICTCOSTM").Columns.Add("SEQ_DESC")
            .Tables("ICTCOSTM").Columns.Add("ITEM_COST_TOTAL_D", GetType(System.Decimal), "ISNULL(ITEM_COST_VCOST,0)+ISNULL(ITEM_COST_LANDG,0)+ISNULL(ITEM_COST_TOOLG,0)+ISNULL(ITEM_COST_OVRHD,0)")
            .Tables("ICTCOSTM").Columns.Add("ITEM_COST_TOTAL_M", GetType(System.Decimal), "ISNULL(ITEM_COST_MATLS,0)+ISNULL(ITEM_COST_LANDGI,0)+ISNULL(ITEM_COST_TOOLGI,0)+ISNULL(ITEM_COST_OVRHDI,0)")


            ASCMAIN1.sql = "Select ICTCOSTH.* from ICTCOSTH where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTH", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ICTCOSTC.* from ICTCOSTC where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTC", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ICTCOSTF.* from ICTCOSTF where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTF", "**", 0, True, "V")


            ASCMAIN1.sql = "Select ICTIREC5.*, GLTACCT1.ACCT_DESC" _
                & " from ICTIREC5,GLTACCT1 where GLTACCT1.ACCT_CODE (+) = ICTIREC5.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIREC5", "**", 2)

            ASCMAIN1.sql = "Select * from ASTAUDT1 where TABLE_NAME='ICTIREC1' AND KEY_VALUE=:PARM1"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from APTACRC1 where RECEIPT_NO = :PARM1"
            Create_TDA(.Tables.Add, "APTACRC1", "**", 0, False, "V")

        End With

        grdICTVARS1.DataSource = dst.Tables("ICTVARS1")
        grdICTCOSTM.DataSource = dst.Tables("ICTCOSTM")
        grdICTIREC5.DataSource = dst.Tables("ICTIREC5")

        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")

        With grdASTAUDTX.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.False
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        End With
        With grdASTAUDTX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With


        With grdICTCOSTM.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_COST_TOTAL", "SEQ_DESC", "OPS_YYYYPP"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Columns("ITEM_COST_TOTAL").CellAppearance.BackColor = Color.LightBlue
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTCOSTM}
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_TOOLG").Hidden = False
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_OVRHD").Hidden = True
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_TOOLGI").Hidden = True
            grd.DisplayLayout.Bands(0).Columns("ITEM_COST_OVRHDI").Hidden = True
        Next


        Set_SEGS(grdICTIREC5, "ICTIREC5")
        With grdICTIREC5.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        For Each C As String In New String() {"ITEM_CODE"}
            grdICTVARS1.DisplayLayout.Bands(0).Columns(C).Header.Fixed = True
        Next
        For Each C As String In New String() {"PO_STATUS", "COST_CATGY_CODE", "OPS_YYYYPP", "BM_ISSUE_NO", "FRT_CLASS_CODE", "TRF_CLASS_CODE", "CFRT_CLASS", "CTRF_CLASS", "FFRT_CLASS", "FTRF_CLASS"}
            grdICTVARS1.DisplayLayout.Bands(0).Columns(C).Header.Appearance.TextHAlign = HAlign.Center
            grdICTVARS1.DisplayLayout.Bands(0).Columns(C).CellAppearance.TextHAlign = HAlign.Center
        Next
        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTVARS1.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            If New String() {"ITEM_CODE", "ITEM_DESC", "PROD_CODE", "COST_CATGY_CODE"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            ElseIf New String() {"CVCOST", "CLANDG", "CTOOLG", "CFRT_CLASS", "CTRF_CLASS", "CFRT_PCT", "CTRF_PCT"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LimeGreen
            ElseIf New String() {"FVCOST", "FLANDG", "FTOOLG", "FFRT_CLASS", "FTRF_CLASS", "FFRT_PCT", "FTRF_PCT"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.DodgerBlue
            ElseIf New String() {"VEND_CODE", "VEND_COUNTRY", "FRT_CLASS_CODE", "TRF_CLASS_CODE"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf New String() {"PO_COST", "PO_COST_FRT", "PO_COST_TRF", "FRT_CLASS_PCT_CUR", "TRF_CLASS_PCT_CUR"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_QTY_ORD", "PO_QTY_OPN", "PO_STATUS"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf New String() {"BM_ISSUE_NO", "BM_ISSUE_SEL"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
            ElseIf New String() {"RECEIPT_NO", "RECEIPT_LNO", "QTY_REC", "QTY_INV", "OPS_YYYYPP", "ACCRUAL_STATUS"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
            ElseIf New String() {"TRAN_PV", "TRAN_MV", "TRAN_FV", "TRAN_TV"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightSteelBlue
                gcol.CellAppearance.BackColor = System.Drawing.Color.LightBlue
            End If

        Next


        Create_Summary(grdICTVARS1, "ITEM_CODE", "Count")
        Create_Summary(grdICTVARS1, New String() {"TRAN_PV", "TRAN_MV", "TRAN_FV", "TRAN_TV"})

        Show_Filter(grdICTVARS1, True)


        'Bind_Controls(Absx1.txtFor("TRF_CLASS_CODE"), "POTORDR1")
        'Bind_Controls(Absx1.txtFor("FRT_CLASS_CODE"), "POTORDR1")
        Bind_Controls(grpPOTORDR1, "POTORDR1")
        Bind_Controls(grpICTIREC1, "ICTIREC1")

        ASCMAIN1.Add_Value_List(grdICTCOSTM, "ITEM_COST_MAKE_BUY", Nothing, New String() {":", "M:Make", "B:Buy"})
        ASCMAIN1.Add_Value_List(grdICTIREC5, "DIST_TYPE", , New String() {":", "TOOLG:TARIFF"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                ' EMsg &= vbCr & "Not Yet"

            Case "Update"

                'If dst.Tables("SOTOXFRX").Select("SEL='1'").Length = 0 Then
                '    EMsg &= vbCr & "Nothing Selected to Transfer"
                'End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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

            Case "Refresh"
                Refresh_Documents()

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                'Update_Record()
                'Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Settings.Enabled = iScreenMode
                    '.Items("Refresh").Visible = False

                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Load").Visible = Not tf
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False

                    .Items("Done").Settings.Enabled = iScreenMode
                End With

                .Groups("View Options").Visible = tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode
        SplitContainer2.Visible = Not ScreenMode

        If ScreenMode Then
            'grdICTVARS1.Parent = SplitContainer1.Panel1
            chkEdit.Visible = (ASCMAIN1.USER_SECURITY_CODEs.Contains("IM"))
        Else
            Clear_Record()
            'grdICTVARS1.Parent = SplitContainer2.Panel1

            grdICTCOSTM.Text = "Cost Details"
            grdICTCOSTM.DisplayLayout.Bands(0).Columns("ITEM_CODE").Hidden = True
            Show_Filter(grdICTCOSTM, False)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"ICTVARS1", "ICTCOSTM", "ICTIREC5", "ICTIREC1", "POTORDR1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
        'Absx1.txtFor("WHSE_CODE").Text = ""

        chkEdit.Checked = False
        Set_EditMode()

        Refresh_Documents()

        chkProjectVariances.Checked = True
        chkShowVariancesOnly.Checked = True

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Create_WorkTables(False)
        Fill_Records("ICTVARS1")

        Set_TRAN_Variances(True)
        Sort_grdColumns(grdICTVARS1, "ITEM_CODE")

        'Set_ICTVARS1_Filter()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_ICTCOSTM(ITEM_CODE As String)

        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        Dim ITEM_YYYYPP_PRV_COST As String = rowICTITEM1.Item("ITEM_YYYYPP_PRV_COST") & ""
        Dim ITEM_YYYYPP_CUR_COST As String = rowICTITEM1.Item("ITEM_YYYYPP_CUR_COST") & ""

        Fill_Records("ICTCOSTH", ITEM_CODE)
        Fill_Records("ICTCOSTC", ITEM_CODE)
        Fill_Records("ICTCOSTF", ITEM_CODE)

        'ASCMAIN1.sql = "Select * from ICTCOSTF where ITEM_EXP_IMP_IND = 'E'"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
        '    rowICTCOSTM = Add_ICTCOSTM(2, row)
        'Next
        'Sort_grdColumns(grdICTCOSTM, "INIT_DATE".ToLower)

        dst.Tables("ICTCOSTM").Rows.Clear()
        rowICTCOSTM = Add_ICTCOSTM(ITEM_CODE, 0, dst.Tables("ICTCOSTH").Rows.Find(New String() {ITEM_YYYYPP_PRV_COST, ITEM_CODE}))
        rowICTCOSTM = Add_ICTCOSTM(ITEM_CODE, 1, dst.Tables("ICTCOSTC").Rows.Find(New String() {ITEM_CODE}))
        rowICTCOSTM.Item("OPS_YYYYPP") = ITEM_YYYYPP_CUR_COST
        rowICTCOSTM = Add_ICTCOSTM(ITEM_CODE, 2, dst.Tables("ICTCOSTF").Rows.Find(New String() {ITEM_CODE}))
        Sort_grdColumns(grdICTCOSTM, "SEQ")
        grdICTCOSTM.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        For Each grow As UltraWinGrid.UltraGridRow In grdICTCOSTM.Rows
            If (grow.Cells("SEQ").Value = 2 And (EntryMode = "E")) _
            Or (grow.Cells("SEQ").Value = 1 And (EntryMode = "V")) Then
                grdICTCOSTM.ActiveRow = grow
                Exit For
            End If
        Next

        rowICTCOSTM = dst.Tables("ICTCOSTM").Select("SEQ = 2")(0)
        If rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") & "" = "" Then
            Dim rowBMTMAIN1 As DataRow = LookUp("BMTMAIN1", ITEM_CODE)
            If rowBMTMAIN1 IsNot Nothing Then
                rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "M"
            Else
                rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "B"
            End If
        End If

        If rowICTCOSTM.Item("ITEM_COST_MAKE_BUY") = "M" Then

            ASCMAIN1.sql = "Select Max (BM_ISSUE_NO) from BMTMAIN2 " & vbCrLf _
                & " where BM_PROD_ITEM = '" & ITEM_CODE & "' and BM_ISSUE_USE_FOR_STD = '1' and BM_ISSUE_NO <> '00'"
            Dim BMI As String = ASCDATA1.GetDataValue

            'If rowICTCOSTM.Item("BM_ISSUE_NO") & "" <> BMI Then
            '    blnCalculate = True
            '    rowICTCOSTM.Item("BM_ISSUE_NO") = BMI
            'End If

        End If

        If rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") & "" = "" Then
            rowICTCOSTM.Item("ITEM_COST_FRT_CLASS") = ROWs("ICTPARM1").Item("IC_PARM_FRT_CLASS")
        End If

        grdICTCOSTM.Text = $"Std Cost Data for Item {ITEM_CODE}"

        Dim RECEIPT_NO As String = grdICTVARS1.ActiveRow.Cells("RECEIPT_NO").Value & ""
        Dim RECEIPT_LNO As Int32 = Val(grdICTVARS1.ActiveRow.Cells("RECEIPT_LNO").Value & "")

        If RECEIPT_NO = "" Then
            grdASTAUDTX.Visible = False
        Else
            Fill_Records("ASTAUDTX", RECEIPT_NO)
            grdASTAUDTX.Visible = True
        End If

    End Sub


    Function Add_ICTCOSTM(ITEM_CODE As String, SEQ As Integer, row As DataRow) As DataRow

        Dim rowICTCOSTM As DataRow = dst.Tables("ICTCOSTM").NewRow
        Dim use_CUR_for_FUT As Boolean = False
        If SEQ = 2 And row Is Nothing Then
            row = dst.Tables("ICTCOSTC").Rows(0)
            use_CUR_for_FUT = True
        End If

        If row IsNot Nothing Then
            For Each dcol As DataColumn In row.Table.Columns
                rowICTCOSTM.Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
            Next
            If use_CUR_for_FUT Then
                rowICTCOSTM.Item("INIT_DATE") = DBNull.Value
                rowICTCOSTM.Item("INIT_OPER") = "Default"
            End If
        Else
            rowICTCOSTM.Item("ITEM_CODE") = ITEM_CODE
        End If

        Dim SEQ_DESC As String = ""
        If SEQ = 0 Then SEQ_DESC = "Prv"
        If SEQ = 1 Then SEQ_DESC = "Cur"
        If SEQ = 2 Then SEQ_DESC = "Fut"

        rowICTCOSTM.Item("SEQ") = SEQ
        rowICTCOSTM.Item("SEQ_DESC") = SEQ_DESC

        If rowICTCOSTM.Item("OPS_YYYYPP") & String.Empty = "" Then
            rowICTCOSTM.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        End If

        dst.Tables("ICTCOSTM").Rows.Add(rowICTCOSTM)
        Return rowICTCOSTM
    End Function


    Sub Update_Record(Optional publish As Boolean = False, Optional showCommitMsg As Boolean = True)

        BeginTrans()

        Stop

        CommitTrans("Update Successful")

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

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTVARS1, "SSBBBBBB", "Show Filter", "Show GroupBox", "Item Status Inquiry", "PO Entry", "PO Inquiry", "PO Receipts Inquiry", "Item Cost Maintenance", "Item Cost Inquiry") ' , "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        ' Load_Popup_Menu(grdSOTORDRX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")

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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTVARS1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Item Cost Maintenance"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.tblASTMENU1.Select("MENU_ITEM_OBJECT = 'ICFCOST1'").Length <> 0)
                    tlb_btn = DirectCast(tlb_pop.Tools("Item Cost Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.tblASTMENU1.Select("MENU_ITEM_OBJECT = 'ICFCOSTI'").Length <> 0)
            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Item Cost Maintenance", "Item Cost Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    If e.Tool.Key = "Item Cost Maintenance" Then
                        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFCOST1")
                    Else
                        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFCOSTI")
                    End If
                End If

            Case "PO Entry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDR1")

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "PO Receipts Inquiry"
                Dim RECEIPT_NO As String = grd.ActiveRow.Cells("RECEIPT_NO").Text
                Context_Launch("View", RECEIPT_NO, e.Tool.Key, "ICFIRECI")

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)

        Select Case COLUMN_NAME
            Case "FRT_CLASS_CODE"
                Absx1.txtFor(COLUMN_NAME).Text = Absx1.txtFor(COLUMN_NAME).Text.ToUpper
            Case "TRF_CLASS_CODE"
                Absx1.txtFor(COLUMN_NAME).Text = Absx1.txtFor(COLUMN_NAME).Text.ToUpper
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "FRT_CLASS_CODE"
            Case "TRF_CLASS_CODE"
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "FRT_CLASS_CODE"
            Case "TRF_CLASS_CODE"
        End Select
    End Sub

#End Region

    Sub Refresh_Documents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Order from Records Selected in Transfer Queue", "")

        Create_WorkTables(False)
        Fill_Records("ICTVARS1")
        Set_TRAN_Variances(True)
        Sort_grdColumns(grdICTVARS1, "ITEM_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdICTVARS1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTVARS1.AfterRowActivate
        Setup_grdICTVARS1_ActiveRow()
    End Sub

    Sub Setup_grdICTVARS1_ActiveRow()
        If grdICTVARS1.ActiveRow Is Nothing OrElse Not grdICTVARS1.ActiveRow.IsDataRow OrElse grdICTVARS1.ActiveRow.IsFilteredOut Then
            splItemDetails.Visible = False
        Else
            Setup_ICTIRECX()

            splItemDetails.Visible = True

            chkEdit.Checked = False
        End If
    End Sub

    Sub Setup_ICTIRECX()

        Dim ITEM_CODE As String = grdICTVARS1.ActiveRow.Cells("ITEM_CODE").Value

        Setup_ICTCOSTM(ITEM_CODE)

        Dim RECEIPT_NO As String = grdICTVARS1.ActiveRow.Cells("RECEIPT_NO").Value & ""
        Dim RECEIPT_LNO As Integer = Val(grdICTVARS1.ActiveRow.Cells("RECEIPT_LNO").Value & "")
        Fill_Records("ICTIREC1", New Object() {RECEIPT_NO})
        grpICTIREC1.Visible = (RECEIPT_NO <> "")

        Fill_Records("ICTIREC5", New Object() {RECEIPT_NO, RECEIPT_LNO})
        grdICTIREC5.Visible = (RECEIPT_NO <> "")

        Dim PO_ORDER_NO As String = grdICTVARS1.ActiveRow.Cells("PO_ORDER_NO").Value & ""
        Dim PO_ORDER_LNO As Integer = Val(grdICTVARS1.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
        Dim rowPOTORDR1 As DataRow = Fill_Record("POTORDR1", New Object() {PO_ORDER_NO})

        Dim FRT_CLASS_CODE As String = rowPOTORDR1.Item("FRT_CLASS_CODE") & ""
        txtFRT_CLASS_CODE.Text = FRT_CLASS_CODE
        If FRT_CLASS_CODE = "" Then
            txtFRT_CLASS_DESC.Text = ""
            numFRT_CLASS_PCT.Value = DBNull.Value
        Else
            Dim rowICTFRTC1 As DataRow = dst.Tables("ICTFRTC1").Rows.Find(FRT_CLASS_CODE)
            txtFRT_CLASS_DESC.Text = rowICTFRTC1.Item("FRT_CLASS_DESC") & ""
            numFRT_CLASS_PCT.Value = Val(rowICTFRTC1.Item("FRT_CLASS_PCT_CUR") & "")
        End If

        Dim TRF_CLASS_CODE As String = rowPOTORDR1.Item("TRF_CLASS_CODE") & ""
        txtTRF_CLASS_CODE.Text = TRF_CLASS_CODE
        If TRF_CLASS_CODE = "" Then
            txtTRF_CLASS_DESC.Text = ""
            numTRF_CLASS_PCT.Value = DBNull.Value
        Else
            Dim rowICTTRFC1 As DataRow = dst.Tables("ICTTRFC1").Rows.Find(TRF_CLASS_CODE)
            txtTRF_CLASS_DESC.Text = rowICTTRFC1.Item("TRF_CLASS_DESC") & ""
            numTRF_CLASS_PCT.Value = Val(rowICTTRFC1.Item("TRF_CLASS_PCT_CUR") & "")
        End If

        Absx1.txtFor("FRT_CLASS_CODE").Focus()
        Absx1.txtFor("TRF_CLASS_CODE").Focus()

    End Sub

    Sub Show_GL()

        'If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
        '    optGL.Tag = optGL.Value
        '    If optGL.Value = "A" Then
        '        grdICTIREC3.DataSource = dst.Tables("ICTIREC3")
        '        Dim dvw As DataView = dst.Tables("ICTIREC3").DefaultView
        '        dvw.RowFilter = ""
        '    ElseIf optGL.Value = "L" Then
        '        grdICTIREC3.DataSource = dst.Tables("ICTIREC3")
        '        Dim dvw As DataView = dst.Tables("ICTIREC3").DefaultView
        '        Dim RECEIPT_LNO As Integer = 0
        '        If grdICTIREC2.ActiveRow IsNot Nothing Then
        '            RECEIPT_LNO = Val(grdICTIREC2.ActiveRow.Cells("RECEIPT_LNO").Text)
        '        End If
        '        dvw.RowFilter = "RECEIPT_LNO = " & CStr(RECEIPT_LNO)
        '    ElseIf optGL.Value = "S" Then
        '        Dim tbl As DataTable = dst.Tables("ICTIREC3").Clone
        '        Dim RECEIPT_GNO As Integer = 0
        '        For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
        '        ("ICTIREC3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
        '            Dim DIST_AMT As Decimal = Val(dst.Tables("ICTIREC3").Compute _
        '            ("SUM(DIST_AMT)",
        '             "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'") & "")
        '            Dim row As DataRow = tbl.NewRow
        '            row.Item("RECEIPT_NO") = Absx1.txtFor("RECEIPT_NO").Text
        '            row.Item("RECEIPT_LNO") = 0
        '            RECEIPT_GNO += 1
        '            row.Item("RECEIPT_GNO") = RECEIPT_GNO
        '            row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
        '            row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
        '            row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
        '            row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
        '            row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
        '            row.Item("DIST_AMT") = DIST_AMT
        '            tbl.Rows.Add(row)
        '        Next

        '        grdICTIREC3.DataSource = tbl
        '    End If
        'End If
    End Sub

    Sub Create_WorkTables(initialize As Boolean, Optional RECEIPT_NO As String = "")

        If initialize Then
            ICTVARS1 = ASCMAIN1.Temp_Table(Get_SQL("ICTVARS1") & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"Create Index I_{ICTVARS1}_1 on {ICTVARS1} (ITEM_CODE, RECEIPT_NO, RECEIPT_LNO)")
            ' ASCDATA1.ExecuteSQL($"Alter Table {ICTVARS1} Add Primary Key (RECEIPT_NO, RECEIPT_LNO)")

        Else
            ASCDATA1.ExecuteSQL($"Truncate Table {ICTVARS1}")
            ASCDATA1.ExecuteSQL($"Insert into {ICTVARS1} {Get_SQL("ICTVARS1", RECEIPT_NO)}")

            ASCMAIN1.sql = $"
Begin 
 Declare Cursor C1 is 
  Select Distinct RECEIPT_NO from {ICTVARS1} where RECEIPT_NO is not null; 
 Begin
  For R1 in C1 Loop 
   ICPIREC5(R1.RECEIPT_NO); 
  End Loop; 
 End; 
End;"
            ASCDATA1.ExecuteSQL()

        End If
    End Sub

    Function Get_SQL(TABLE_NAME As String, Optional RECEIPT_NO As String = "") As String

        Dim SQL As String = ""

        Select Case TABLE_NAME

            Case "ICTVARS1"

                SQL = ""

                If RECEIPT_NO = "" Then
                    SQL = "" _
                        & "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, POTORDR2.ITEM_CODE" & vbCrLf _
                        & ", POTORDR2.PO_COST PO_COST" & vbCrLf _
                        & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_STATUS" & vbCrLf _
                        & ", NULL RECEIPT_NO, NULL RECEIPT_LNO, NULL QTY_REC, NULL QTY_INV, NULL OPS_YYYYPP" & vbCrLf _
                        & ", POTORDR2.BM_ISSUE_SEL BM_ISSUE_SEL" & vbCrLf _
                        & ", POTORDR2.BM_ISSUE_NO BM_ISSUE_NO" & vbCrLf _
                        & ", 0 TRAN_PV,0 TRAN_MV,0 TRAN_FV,0 TRAN_TV" & vbCrLf _
                        & ", NULL ACCRUAL_STATUS" & vbCrLf _
                        & ", ROUND (NVL(ICTFRTC1.FRT_CLASS_PCT_CUR * NVL(POTORDR2.PO_COST,0)/100,NVL(ICTCOSTA.ITEM_COST_LANDG,0)),6) PO_COST_FRT" & vbCrLf _
                        & ", ROUND (NVL(ICTTRFC1.TRF_CLASS_PCT_CUR * NVL(POTORDR2.PO_COST,0)/100,NVL(ICTCOSTA.ITEM_COST_LANDG,0)),6) PO_COST_TRF" & vbCrLf _
                        & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                        & ", ICTCOSTA.ITEM_COST_TOTAL, POTORDR1.VEND_CODE, APTVEND1.VEND_COUNTRY" & vbCrLf _
                        & ", POTORDR1.FRT_CLASS_CODE, POTORDR1.TRF_CLASS_CODE" & vbCrLf _
                        & ", ICTFRTC1.FRT_CLASS_PCT_CUR, ICTTRFC1.TRF_CLASS_PCT_CUR" & vbCrLf _
                        & ", ICTCOSTC.ITEM_COST_VCOST CVCOST, ICTCOSTC.ITEM_COST_LANDG CLANDG, ICTCOSTC.ITEM_COST_TOOLG CTOOLG, ICTCOSTC.ITEM_COST_FRT_CLASS CFRT_CLASS, ICTCOSTC.ITEM_COST_TRF_CLASS CTRF_CLASS" & vbCrLf _
                        & ", ICTFRTC1_C.FRT_CLASS_PCT_CUR CFRT_PCT, ICTTRFC1_C.TRF_CLASS_PCT_CUR CTRF_PCT" & vbCrLf _
                        & ", NVL(ICTCOSTF.ITEM_COST_VCOST,ICTCOSTC.ITEM_COST_VCOST) FVCOST, NVL(ICTCOSTF.ITEM_COST_LANDG,ICTCOSTC.ITEM_COST_LANDG) FLANDG, NVL(ICTCOSTF.ITEM_COST_TOOLG,ICTCOSTC.ITEM_COST_TOOLG) FTOOLG, NVL(ICTCOSTF.ITEM_COST_FRT_CLASS,ICTCOSTC.ITEM_COST_FRT_CLASS) FFRT_CLASS, NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTC.ITEM_COST_TRF_CLASS) FTRF_CLASS" & vbCrLf _
                        & ", ICTFRTC1_F.FRT_CLASS_PCT_CUR FFRT_PCT, ICTTRFC1_F.TRF_CLASS_PCT_CUR FTRF_PCT" & vbCrLf _
                        & "from ICTITEM1,ICTCOSTC,ICTCOSTF,APTVEND1,POTORDR1,POTORDR2,ICTCOSTA" & vbCrLf _
                        & ",ICTTRFC1,ICTFRTC1,ICTFRTC1 ICTFRTC1_C, ICTFRTC1 ICTFRTC1_F,ICTTRFC1 ICTTRFC1_C, ICTTRFC1 ICTTRFC1_F" & vbCrLf _
                        & "where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                        & "  And (POTORDR2.PO_STATUS = 'O' and POTORDR2.PO_QTY_OPN > 0)" & vbCrLf _
                        & "  and ICTCOSTC.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                        & $"  And ICTCOSTA.OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                        & "  and ICTCOSTA.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                        & "  and ICTCOSTF.ITEM_CODE (+) = POTORDR2.ITEM_CODE" & vbCrLf _
                        & "  and ICTFRTC1.FRT_CLASS_CODE (+) = POTORDR1.FRT_CLASS_CODE" & vbCrLf _
                        & "  and ICTTRFC1.TRF_CLASS_CODE (+) = POTORDR1.TRF_CLASS_CODE" & vbCrLf _
                        & "  and ICTFRTC1_C.FRT_CLASS_CODE (+) = ICTCOSTC.ITEM_COST_FRT_CLASS" & vbCrLf _
                        & "  and ICTTRFC1_C.TRF_CLASS_CODE (+) = ICTCOSTC.ITEM_COST_TRF_CLASS" & vbCrLf _
                        & "  and ICTFRTC1_F.FRT_CLASS_CODE (+) = NVL(ICTCOSTF.ITEM_COST_FRT_CLASS,ICTCOSTC.ITEM_COST_FRT_CLASS)" & vbCrLf _
                        & "  and ICTTRFC1_F.TRF_CLASS_CODE (+) = NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTF.ITEM_COST_TRF_CLASS)" & vbCrLf _
                        & "  and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                        & "  and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE" & vbCrLf _
                        & "UNION" & vbCrLf
                End If

                Dim sqlRECEIPT_NO As String = ""
                If RECEIPT_NO <> "" Then
                    sqlRECEIPT_NO = $"  and ICTIREC2.RECEIPT_NO = '{RECEIPT_NO}'"
                End If

                SQL &= "" _
                    & "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO, POTORDR2.ITEM_CODE" & vbCrLf _
                    & ", CASE WHEN ICTIREC2.RECEIPT_NO IS NULL THEN POTORDR2.PO_COST ELSE ICTIREC2.PO_COST END PO_COST" & vbCrLf _
                    & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_STATUS" & vbCrLf _
                    & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC2.QTY_REC, ICTIREC2.QTY_INV, ICTIREC2.OPS_YYYYPP" & vbCrLf _
                    & ", NVL(ICTIREC2.BM_ISSUE_SEL,POTORDR2.BM_ISSUE_SEL) BM_ISSUE_SEL" & vbCrLf _
                    & ", NVL(ICTIREC2.BM_ISSUE_NO,POTORDR2.BM_ISSUE_NO) BM_ISSUE_NO" & vbCrLf _
                    & ", ICTIREC2.TRAN_PV,ICTIREC2.TRAN_MV,ICTIREC2.TRAN_FV,ICTIREC2.TRAN_TV" & vbCrLf _
                    & ", ICTIREC2.ACCRUAL_STATUS" & vbCrLf _
                    & ", ICTIREC2.PO_COST_FRT" & vbCrLf _
                    & ", ICTIREC2.PO_COST_TRF" & vbCrLf _
                    & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                    & ", ICTCOSTA.ITEM_COST_TOTAL, POTORDR1.VEND_CODE, APTVEND1.VEND_COUNTRY" & vbCrLf _
                    & ", POTORDR1.FRT_CLASS_CODE, POTORDR1.TRF_CLASS_CODE" & vbCrLf _
                    & ", ICTFRTC1.FRT_CLASS_PCT_CUR, ICTTRFC1.TRF_CLASS_PCT_CUR" & vbCrLf _
                    & ", ICTCOSTC.ITEM_COST_VCOST CVCOST, ICTCOSTC.ITEM_COST_LANDG CLANDG, ICTCOSTC.ITEM_COST_TOOLG CTOOLG, ICTCOSTC.ITEM_COST_FRT_CLASS CFRT_CLASS, ICTCOSTC.ITEM_COST_TRF_CLASS CTRF_CLASS" & vbCrLf _
                    & ", ICTFRTC1_C.FRT_CLASS_PCT_CUR CFRT_PCT, ICTTRFC1_C.TRF_CLASS_PCT_CUR CTRF_PCT" & vbCrLf _
                    & ", NVL(ICTCOSTF.ITEM_COST_VCOST,ICTCOSTC.ITEM_COST_VCOST) FVCOST, NVL(ICTCOSTF.ITEM_COST_LANDG,ICTCOSTC.ITEM_COST_LANDG) FLANDG, NVL(ICTCOSTF.ITEM_COST_TOOLG,ICTCOSTC.ITEM_COST_TOOLG) FTOOLG, NVL(ICTCOSTF.ITEM_COST_FRT_CLASS,ICTCOSTC.ITEM_COST_FRT_CLASS) FFRT_CLASS, NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTC.ITEM_COST_TRF_CLASS) FTRF_CLASS" & vbCrLf _
                    & ", ICTFRTC1_F.FRT_CLASS_PCT_CUR FFRT_PCT, ICTTRFC1_F.TRF_CLASS_PCT_CUR FTRF_PCT" & vbCrLf _
                    & "from ICTITEM1,ICTCOSTC,ICTCOSTF,ICTIREC2,ICTIREC1,APTVEND1,POTORDR1,POTORDR2,ICTCOSTA" & vbCrLf _
                    & ",ICTTRFC1,ICTFRTC1,ICTFRTC1 ICTFRTC1_C, ICTFRTC1 ICTFRTC1_F,ICTTRFC1 ICTTRFC1_C, ICTTRFC1 ICTTRFC1_F" & vbCrLf _
                    & "where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                    & "  and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                    & $"  And ICTIREC2.OPS_YYYYPP = '{ASCMAIN1.CYP}'" & sqlRECEIPT_NO & vbCrLf _
                    & "  and ICTIREC2.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "  and ICTIREC2.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & "  and ICTCOSTC.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                    & $"  And ICTCOSTA.OPS_YYYYPP = '{ASCMAIN1.CYP}'" & vbCrLf _
                    & "  and ICTCOSTA.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
                    & "  and ICTCOSTF.ITEM_CODE (+) = POTORDR2.ITEM_CODE" & vbCrLf _
                    & "  and ICTFRTC1.FRT_CLASS_CODE (+) = POTORDR1.FRT_CLASS_CODE" & vbCrLf _
                    & "  and ICTTRFC1.TRF_CLASS_CODE (+) = POTORDR1.TRF_CLASS_CODE" & vbCrLf _
                    & "  and ICTFRTC1_C.FRT_CLASS_CODE (+) = ICTCOSTC.ITEM_COST_FRT_CLASS" & vbCrLf _
                    & "  and ICTTRFC1_C.TRF_CLASS_CODE (+) = ICTCOSTC.ITEM_COST_TRF_CLASS" & vbCrLf _
                    & "  and ICTFRTC1_F.FRT_CLASS_CODE (+) = NVL(ICTCOSTF.ITEM_COST_FRT_CLASS,ICTCOSTC.ITEM_COST_FRT_CLASS)" & vbCrLf _
                    & "  and ICTTRFC1_F.TRF_CLASS_CODE (+) = NVL(ICTCOSTF.ITEM_COST_TRF_CLASS,ICTCOSTF.ITEM_COST_TRF_CLASS)" & vbCrLf _
                    & "  and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                    & "  and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"

        End Select

        Return SQL

    End Function

    Private Sub optStatus_ValueChanged(sender As Object, e As EventArgs) Handles optStatus.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_ICTVARS1_Filter()
    End Sub

    Private Sub chkNONUS_CheckedChanged(sender As Object, e As EventArgs) Handles chkNONUS.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_ICTVARS1_Filter()
    End Sub

    Sub Set_ICTVARS1_Filter()
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim SQLW As String = ""
        If chkNONUS.Checked Then SQLW &= " and ISNULL(VEND_COUNTRY,'X') <> 'US' and ISNULL(VEND_COUNTRY,'X') <> 'USA'"
        If optStatus.Value = "O" Then
            SQLW &= " and ISNULL(RECEIPT_NO,'X') = 'X'"
        ElseIf optStatus.Value = "R" Then
            SQLW &= " and ISNULL(RECEIPT_NO,'X') <> 'X'"
        End If

        If chkShowVariancesOnly.Checked Then
            SQLW &= " and (ISNULL(TRAN_PV,0) <> 0 OR ISNULL(TRAN_MV,0) <> 0 OR ISNULL(TRAN_FV,0) <> 0 OR ISNULL(TRAN_TV,0) <> 0)"
        End If

        Dim dvw As DataView = DirectCast(grdICTVARS1.DataSource, DataTable).DefaultView
        If SQLW <> "" Then
            SQLW = Mid(SQLW, 6)
        End If
        dvw.RowFilter = SQLW

        Setup_grdICTVARS1_ActiveRow()

    End Sub

    Sub Set_TRAN_Variances(Optional initialize As Boolean = False, Optional RECEIPT_NO_in As String = "")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Variances")

        Dim sqlRECEIPT_NO As String = ""
        If RECEIPT_NO_in <> "" Then
            sqlRECEIPT_NO = $"RECEIPT_NO = '{RECEIPT_NO_in}'"
        End If
        For Each rowICTVARS1 As DataRow In dst.Tables("ICTVARS1").Select(sqlRECEIPT_NO)

            Dim RECEIPT_NO As String = rowICTVARS1.Item("RECEIPT_NO") & ""
            Dim ITEM_CODE As String = rowICTVARS1.Item("ITEM_CODE") & ""

            ' If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH024C10USA" Then Stop

            If initialize And RECEIPT_NO <> "" Then
                rowICTVARS1.Item("TRAN_PV_ORIG") = rowICTVARS1.Item("TRAN_PV")
                rowICTVARS1.Item("TRAN_MV_ORIG") = rowICTVARS1.Item("TRAN_MV")
                rowICTVARS1.Item("TRAN_FV_ORIG") = rowICTVARS1.Item("TRAN_FV")
                rowICTVARS1.Item("TRAN_TV_ORIG") = rowICTVARS1.Item("TRAN_TV")
            End If

            Calculations(rowICTVARS1)
        Next

        If RECEIPT_NO_in = "" Then
            Set_ICTVARS1_Filter()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Calculations(rowICTVARS1 As DataRow)

        Dim FRT_CLASS_CODE As String = rowICTVARS1.Item("FRT_CLASS_CODE") & ""
        Dim TRF_CLASS_CODE As String = rowICTVARS1.Item("TRF_CLASS_CODE") & ""

        Dim QTY_REC As Decimal = Val(rowICTVARS1.Item("QTY_REC") & "")
        Dim PO_QTY_OPN As Decimal = Val(rowICTVARS1.Item("PO_QTY_OPN") & "")
        If QTY_REC = 0 Then QTY_REC = PO_QTY_OPN

        Dim PO_COST As Decimal = Val(rowICTVARS1.Item("PO_COST") & "")

        Dim CF As String = ""

        If chkProjectVariances.Checked Then
            CF = "FUT"
        Else
            CF = "CUR"
        End If

        ' SHOULD NOT BE NECESSARY AFTER WE FIX THE ICTIREC1/2 DATA

        'Dim rowICTFRTC1 As DataRow = Nothing
        'If FRT_CLASS_CODE <> "" Then rowICTFRTC1 = dst.Tables("ICTFRTC1").Rows.Find(FRT_CLASS_CODE)
        'Dim FRT_CLASS_PCT As Decimal = 0
        'If rowICTFRTC1 IsNot Nothing Then FRT_CLASS_PCT = Val(rowICTFRTC1.Item("FRT_CLASS_PCT_" & CF) & "")
        'rowICTVARS1.Item("PO_COST_FRT") = PO_COST * FRT_CLASS_PCT / 100

        'Dim rowICTTRFC1 As DataRow = Nothing
        'If TRF_CLASS_CODE <> "" Then rowICTTRFC1 = dst.Tables("ICTTRFC1").Rows.Find(TRF_CLASS_CODE)
        'Dim TRF_CLASS_PCT As Decimal = 0
        'If rowICTTRFC1 IsNot Nothing Then TRF_CLASS_PCT = Val(rowICTTRFC1.Item("TRF_CLASS_PCT_" & CF) & "")
        'rowICTVARS1.Item("PO_COST_TRF") = PO_COST * TRF_CLASS_PCT / 100
        'If rowICTVARS1.Item("ITEM_CODE") = "CC001A01" Then Stop
        If CF = "FUT" Then
            Dim PO_COST_FRT As Decimal = Val(rowICTVARS1.Item("PO_COST_FRT") & "")
            Dim PO_COST_TRF As Decimal = Val(rowICTVARS1.Item("PO_COST_TRF") & "")

            Dim FVCOST As Decimal = Val(rowICTVARS1.Item("FVCOST") & "")
            'Dim FMATLS As Decimal = Val(rowICTVARS1.Item("FMATLS") & "")
            Dim FLANDG As Decimal = Val(rowICTVARS1.Item("FLANDG") & "")
            Dim FTOOLG As Decimal = Val(rowICTVARS1.Item("FTOOLG") & "")

            Dim TRAN_PV As Decimal = System.Math.Round(QTY_REC * (PO_COST - FVCOST), 2)
            'Dim TRAN_MV As Decimal = System.Math.Round(QTY_REC * (PO_COST_MTL - FMATLS), 2)
            Dim TRAN_FV As Decimal = System.Math.Round(QTY_REC * (PO_COST_FRT - FLANDG), 2)
            Dim TRAN_TV As Decimal = System.Math.Round(QTY_REC * (PO_COST_TRF - FTOOLG), 2)

            rowICTVARS1.Item("TRAN_PV") = TRAN_PV
            'rowICTVARS1.Item("TRAN_MV") = TRAN_MV
            rowICTVARS1.Item("TRAN_FV") = TRAN_FV
            rowICTVARS1.Item("TRAN_TV") = TRAN_TV
        Else
            rowICTVARS1.Item("TRAN_PV") = rowICTVARS1.Item("TRAN_PV_ORIG")
            'rowICTVARS1.Item("TRAN_MV") = rowICTVARS1.Item("TRAN_MV_ORIG")
            rowICTVARS1.Item("TRAN_FV") = rowICTVARS1.Item("TRAN_FV_ORIG")
            rowICTVARS1.Item("TRAN_TV") = rowICTVARS1.Item("TRAN_TV_ORIG")
        End If


    End Sub
    Private Sub chkProjectVariances_CheckedChanged(sender As Object, e As EventArgs) Handles chkProjectVariances.CheckedChanged
        Set_TRAN_Variances()
    End Sub

    Private Sub chkShowVariancesOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowVariancesOnly.CheckedChanged
        Set_ICTVARS1_Filter()
    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim FRT_CLASS_CODE As String = Absx1.txtFor("FRT_CLASS_CODE").Text
        Dim TRF_CLASS_CODE As String = Absx1.txtFor("TRF_CLASS_CODE").Text
        Dim FRT_CLASS_DESC As String = Absx1.txtFor("FRT_CLASS_DESC").Text
        Dim TRF_CLASS_DESC As String = Absx1.txtFor("TRF_CLASS_DESC").Text

        If FRT_CLASS_CODE = "" Or TRF_CLASS_CODE = "" Or FRT_CLASS_DESC = "" Or TRF_CLASS_DESC = "" Then
            MsgBox("Invalid or Missing Code - Freight & Tariff Codes are Mandatory", MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        Dim RECEIPT_NO As String = grdICTVARS1.ActiveRow.Cells("RECEIPT_NO").Value
        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)

        Dim FRT_CLASS_PCT As Decimal = Absx1.numFor("FRT_CLASS_PCT_CUR").Value
        Dim TRF_CLASS_PCT As Decimal = Absx1.numFor("TRF_CLASS_PCT_CUR").Value

        rowICTIREC1.Item("FRT_CLASS_PCT") = FRT_CLASS_PCT
        rowICTIREC1.Item("TRF_CLASS_PCT") = TRF_CLASS_PCT

        Try
            BeginTrans()

            Update_Record_TDA("ICTIREC1")

            Dim sqlw As String = $"where RECEIPT_NO = '{RECEIPT_NO}'"
            Dim sqlQTY_REC As String = "(Select QTY_REC * PO_COST from ICTIREC2 where RECEIPT_NO = APTACRC1.RECEIPT_NO and RECEIPT_LNO = APTACRC1.RECEIPT_LNO)"

            ASCMAIN1.sql = $"Update APTACRC1 Set XXX_CLASS_CODE = '{FRT_CLASS_CODE}', XXX_CLASS_PCT = {FRT_CLASS_PCT}, 
  COST_ACC = Round({FRT_CLASS_PCT} / 100 * {sqlQTY_REC},2) {sqlw} and ACCRUAL_CODE = 'FRT'"
            Dim rFRT As Integer = ASCDATA1.ExecuteSQL()
            If rFRT = 0 Then
                TAC.ICCMAIN1.Create_Accrual_FRT(RECEIPT_NO)
            End If

            ASCMAIN1.sql = $"Update APTACRC1 Set XXX_CLASS_CODE = '{TRF_CLASS_CODE}', XXX_CLASS_PCT = {TRF_CLASS_PCT}, 
  COST_ACC = Round({TRF_CLASS_PCT} / 100 * {sqlQTY_REC},2) {sqlw} and ACCRUAL_CODE = 'TRF'"
            Dim rTRF As Integer = ASCDATA1.ExecuteSQL()
            If rTRF = 0 Then
                TAC.ICCMAIN1.Create_Accrual_TRF(RECEIPT_NO)
            End If

            ASCMAIN1.sql = $"Update ICTIREC2 Set PO_COST_FRT = Round({FRT_CLASS_PCT} * PO_COST / 100, 6) {sqlw}"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Update ICTIREC2 Set PO_COST_TRF = Round({TRF_CLASS_PCT} * PO_COST / 100, 6) {sqlw}"
            ASCDATA1.ExecuteSQL()

            ' MAYBE WE SHOULD INCORPORATE ICPIREC5 INTO ICPIRECV
            ASCMAIN1.sql = $"
Begin 
   ICPIRECV({RECEIPT_NO});
   ICPIREC5({RECEIPT_NO});
End;"
            ASCDATA1.ExecuteSQL()

            CommitTrans()
        Catch ex As Exception

            MsgBox("Error - " & ex.InnerException.Message, vbOKOnly, "Update was not successful")
            Rollback()

        End Try

        Create_WorkTables(False, RECEIPT_NO)
        ASCMAIN1.sql = $"Select * from {ICTVARS1}"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        For Each row As DataRow In tbl.Select()
            Dim RECEIPT_LNO As Int32 = Val(row.Item("RECEIPT_LNO") & "")
            Dim rowICTVARS1 As DataRow = dst.Tables("ICTVARS1").Select($"RECEIPT_NO = '{RECEIPT_NO}' and RECEIPT_LNO = {RECEIPT_LNO}")(0)
            rowICTVARS1.ItemArray = row.ItemArray
        Next

        Set_TRAN_Variances(True, RECEIPT_NO)

        Setup_grdICTVARS1_ActiveRow()


        chkEdit.Checked = False
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        If Me.SELECTION_NO = 0 Then Exit Sub

        chkEdit.Checked = False
    End Sub

    Private Sub chkEdit_CheckedChanged(sender As Object, e As EventArgs) Handles chkEdit.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_EditMode()
    End Sub

    Sub Set_EditMode()

        Dim FRT_invoiced As Boolean = False
        Dim TRF_invoiced As Boolean = False

        If chkEdit.Checked Then
            Dim PO_ORDER_NO As String = grdICTVARS1.ActiveRow.Cells("PO_ORDER_NO").Value
            If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO,,,, 1) Then
                chkEdit.Checked = False
                Exit Sub
            End If
            Dim VEND_CODE As String = grdICTVARS1.ActiveRow.Cells("VEND_CODE").Value
            If Not ASCMAIN1.Logical_Lock("APTVEND1", VEND_CODE,,,, 1) Then
                chkEdit.Checked = False
                Exit Sub
            End If

            If Not ASCMAIN1.Logical_Lock("R", "APRACRX1",,,, 1) Then
                chkEdit.Checked = False
                Exit Sub
            End If

            Dim RECEIPT_NO As String = grdICTVARS1.ActiveRow.Cells("RECEIPT_NO").Value & ""
            If RECEIPT_NO <> "" Then
                Fill_Records("APTACRC1", New Object() {RECEIPT_NO})
                FRT_invoiced = (dst.Tables("APTACRC1").Select("ACCRUAL_CODE = 'FRT' and CTL_STATUS = '1'").Length > 0)
                TRF_invoiced = (dst.Tables("APTACRC1").Select("ACCRUAL_CODE = 'TRF' and CTL_STATUS = '1'").Length > 0)
                If FRT_invoiced And TRF_invoiced Then 'dst.Tables("APTACRC1").Select("CTL_STATUS = '1'").Length > 0 Then
                    MsgBox("All Accruals for This PO Receipt have been Invoiced", MsgBoxStyle.OkOnly, "Cannot Change Receipts Data")
                    chkEdit.Checked = False
                    Exit Sub
                End If
            End If
        Else
            ASCMAIN1.MultiTask_Release(,, 1)
            dst.Tables("ICTIREC1").RejectChanges()
        End If


        cmdUpdate.Visible = chkEdit.Checked
        cmdCancel.Visible = chkEdit.Checked
        chkEdit.Visible = Not chkEdit.Checked

        Absx1.txtFor("FRT_CLASS_CODE").ReadOnly = Not chkEdit.Checked Or FRT_invoiced
        Absx1.txtFor("TRF_CLASS_CODE").ReadOnly = Not chkEdit.Checked Or TRF_invoiced

        If grdICTVARS1.ActiveRow IsNot Nothing Then
            Setup_ICTIRECX()

            'Dim ITEM_CODE As String = grdICTVARS1.ActiveRow.Cells("ITEM_CODE").Value
            'Setup_ICTCOSTM(ITEM_CODE)
        End If
    End Sub
End Class