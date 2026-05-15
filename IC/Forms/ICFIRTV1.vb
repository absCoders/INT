Imports System.Drawing
Imports System.Math


Public Class ICFIRTV1

    Dim rowICTIRTV1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIRTVI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTIRTV1.*" _
            & " from ICTIRTV1 where ICTIRTV1.OPS_YYYYPP = :PARM1"
            Call Create_TDA(.Tables.Add, "ICTIRTVX", "**", 0, False, "V")

            Call Create_TDA(.Tables.Add, "ICTIRTV1", "*")

            ASCMAIN1.sql = "Select IR2.*, II1.ITEM_DESC, NVL(IS2.WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND" & vbCrLf _
            & " from ICTIRTV2 IR2 JOIN ICTITEM1 II1 ON (II1.ITEM_CODE = IR2.ITEM_CODE)" & vbCrLf _
            & " LEFT JOIN ICTSTAT2 IS2 ON (IS2.ITEM_CODE=IR2.ITEM_CODE AND IS2.WHSE_CODE=:PARM1)" & vbCrLf _
            & " WHERE IR2.RTV_NO=:PARM2"
            Call Create_TDA(.Tables.Add, "ICTIRTV2", "**", 0, True, "VV")
            .Tables("ICTIRTV2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(RTV_QTY,0) * ISNULL(ITEM_COST_STD,0)")
            .Tables("ICTIRTV2").Columns.Add("PO_LINE_COSTS", GetType(System.Decimal), "ISNULL(RTV_QTY,0) * ISNULL(PO_COST,0)")
 
            ASCMAIN1.sql = "Select ICTIRTV3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIRTV3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIRTV3.ACCT_CODE"
            Call Create_TDA(.Tables.Add, "ICTIRTV3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Call Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select ICTWHSE1.*" _
            & " from ICTWHSE1 where WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "V")


            .Tables.Add("ICTIRTV0")
            .Tables("ICTIRTV0").Columns.Add("KEY")
            .Tables("ICTIRTV0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTCATG1"
            Call Create_TDA(.Tables.Add, "ICTCATG1", "**", 0, False)

            ASCMAIN1.sql = "Select * from APTVEND1 where VEND_CODE = :PARM1"
            Call Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTINVH1", "*")
            Create_TDA(.Tables.Add, "APTINVH2", "*")

            ASCMAIN1.sql = "SELECT ICTITEM1.* FROM ICTITEM1 " _
            & " WHERE ITEM_CODE = :PARM1 OR ITEM_UPC_CODE = :PARM1 OR ITEM_EAN_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEMS", "**", 0, False, "V")

            ASCMAIN1.sql = "SELECT * FROM ICTITEM1 " _
            & " WHERE ITEM_CODE = :PARM1 OR ITEM_UPC_CODE = :PARM1 OR ITEM_EAN_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "V")

        End With

        '   Set_Read_Only(grpTotals, True)

        Fill_Records("ICTCATG1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)


        grdICTIRTV0.DataSource = dst.Tables("ICTIRTV0")
        grdICTIRTV2.DataSource = dst.Tables("ICTIRTV2")
        grdICTIRTV3.DataSource = dst.Tables("ICTIRTV3")
        grdICTIRTVX.DataSource = dst.Tables("ICTIRTVX")

        Call Create_Summary(grdICTIRTVX, "RTV_NO", "Count")
        Call Create_Summary(grdICTIRTVX, "TOTAL_COSTS")

        Call Create_Summary(grdICTIRTV2, "RTV_LNO", "Count")
        Call Create_Summary(grdICTIRTV2, "RTV_QTY")
        Call Create_Summary(grdICTIRTV2, "LINE_COSTS")
        Call Create_Summary(grdICTIRTV2, "PO_LINE_COSTS")

        Call Create_Summary(grdICTIRTV3, "RTV_GNO", "Count")
        Call Create_Summary(grdICTIRTV3, "DIST_AMT")

        Dim rtvStatus As New Infragistics.Win.ValueList
        rtvStatus.ValueListItems.Add("F", "Finalized")
        rtvStatus.ValueListItems.Add("H", "On Hold")
        grdICTIRTVX.DisplayLayout.Bands(0).Columns("RTV_STATUS").ValueList = rtvStatus


        With grdICTIRTVX.DisplayLayout.Bands("ICTIRTVX")
            .Columns("RTV_NO").Header.Fixed = True
        End With

        grdICTIRTV0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Call Set_SEGS(grdICTIRTV3, "ICTIRTV3")

        ' Set_Read_Only(grpTotals, True)
        'If InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0 Then
        '    grpTotals.Visible = False
        '    With grdICTIRTV2.DisplayLayout.Bands(0)
        '        .Columns("ITEM_COST_STD").Hidden = True
        '        .Columns("LINE_COSTS").Hidden = True
        '        '.Columns("COST_CATGY_CODE").Hidden = True
        '        .Columns("PROD_CODE").Hidden = True
        '    End With
        'End If

        grpHeader.Visible = False

        Check_InquiryMode()

    End Sub

    Sub Check_InquiryMode()
        If InquiryMode Then
            With UltraExplorerBar1.Groups("Screen Control")
                .Items("New").Visible = False
                .Items("Update").Visible = False
                .Items("Cancel").Visible = False
            End With
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("VEND_CODE")
                Validate_Code("WHSE_CODE")

                If Absx1.dteFor("RTV_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If

            Case "View"
                If Absx1.txtFor("RTV_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIRTV1 = LookUp("ICTIRTV1", Absx1.txtFor("RTV_NO").Text)
                    If rowICTIRTV1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("RTV_NO").Text & " on File"
                    Else
                        If Not ASCMAIN1.Logical_Lock("ICTIRTV1", Absx1.txtFor("RTV_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Reverse RTV"
                If MessageBox.Show("Are you sure you want to reverse this RTV?", "Confirm Reversal", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, 0, False) = Windows.Forms.DialogResult.No Then
                    EMsg &= vbCr & "Reversal canceled."
                End If

            Case "Update"
                Validate_Code("WH_OPER_ID")

                'If Absx1.txtFor("WH_OPER_ID").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify an Operator ID"
                'End If

                If grdICTIRTV2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTIRTV2 As DataRow In dst.Tables("ICTIRTV2").Select("", "", DataViewRowState.CurrentRows)
                        If rowICTIRTV2.Item("PRICE_CATGY_CODE") & "" = "" Then
                            EMsg &= "Unable to determine Price Category Code for " & rowICTIRTV2.Item("ITEM_CODE") & " (See Line " & rowICTIRTV2.Item("RTV_LNO") & ")"
                        End If
                        If rowICTIRTV2.Item("PROD_CATGY_CODE") & "" = "" Then
                            EMsg &= "Unable to determine Product Category Code for " & rowICTIRTV2.Item("ITEM_CODE") & " (See Line " & rowICTIRTV2.Item("RTV_LNO") & ")"
                        End If
                        If Val(rowICTIRTV2.Item("RTV_QTY") & "") <= 0 Then
                            EMsg &= "Invalid RTV Qty for " & rowICTIRTV2.Item("ITEM_CODE") & " (See Line " & rowICTIRTV2.Item("RTV_LNO") & ")"
                        End If

                    Next

                End If

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIRTV2", dst.Tables("ICTIRTV1").Rows(0).Item("WHSE_CODE"), "RTV_QTY", -1)
                    If msg <> "" Then

                        If MsgBox(msg & vbCr & vbCr & "OK to Continue Anyway?", MsgBoxStyle.YesNo, "The following Items do not have Sufficent Qty for this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Compare Inventory"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Reverse RTV"
                Update_Record(True)
                Call Mode_Settings(False)

            Case "Update"

                Call Update_Record()
                Print_RTV()

                Call Mode_Settings(False)

            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Print RTV"
                Print_RTV()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1

                If InquiryMode Or Not ScreenMode Then
                    .Groups("Screen Control").Items("Reverse RTV").Visible = False
                ElseIf Not InquiryMode And ScreenMode Then
                    .Groups("Screen Control").Items("Reverse RTV").Visible = True
                End If

                .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode

                '   .Groups("Screen Control").Items("Compare Inventory").Visible = ScreenMode

                If ScreenMode And EntryMode <> "N" Then
                    .Groups("Screen Control").Items("Update").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                End If

                If ScreenMode And EntryMode <> "V" Then
                    .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Print RTV").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Reverse RTV").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Print RTV").Settings.Enabled = iScreenMode
                    If ASCMAIN1.USER_SECURITY_CODEs.Contains("WM") AndAlso rowICTIRTV1 IsNot Nothing AndAlso rowICTIRTV1.Item("REVERSES_RTV_NO") Is DBNull.Value AndAlso rowICTIRTV1.Item("REVERSED_BY_RTV_NO") Is DBNull.Value Then
                        .Groups("Screen Control").Items("Reverse RTV").Settings.Enabled = iScreenMode
                    End If
                End If

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") <> 0
                .Groups("Display Options").Visible = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode
        grdICTIRTVX.Visible = Not ScreenMode

        If ScreenMode Then
            grdICTIRTV2.Visible = True
            grdICTIRTV0.Visible = (EntryMode = "V")
            SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                With grdICTIRTV2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
                With grdICTIRTV2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("RTV_QTY").CellAppearance.BackColor = Color.LightYellow
                End With
                With grdICTIRTV3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            Else
                With grdICTIRTV2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
                With grdICTIRTV2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("RTV_QTY").CellAppearance.BackColor = Color.Empty
                End With
                With grdICTIRTV3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ICTIRTV0").Rows.Clear()
        dst.Tables("ICTIRTV1").Rows.Clear()
        dst.Tables("ICTIRTV2").Rows.Clear()
        dst.Tables("ICTIRTV3").Rows.Clear()
        dst.EnforceConstraints = True

        Refresh_Documents()

        Absx1.txtFor("VEND_CODE").Text = ""
        Absx1.dteFor("RTV_DATE").Value = Format(Now, "MM/dd/yyyy")
        ASCMAIN1.MultiTask_Release()
        Absx1.txtFor("RTV_NO").Text = ""

        optGL.Tag = ""
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTIRTV1 = dst.Tables("ICTIRTV1").NewRow
            rowICTIRTV1.Item("RTV_NO") = ASCMAIN1.Next_Control_No("ICTIRTV1.RTV_NO")
            rowICTIRTV1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowICTIRTV1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTIRTV1.Item("RTV_DATE") = HFs("RTV_DATE")
            rowICTIRTV1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIRTV1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIRTV1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIRTV1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTIRTV1.Item("LAST_DATE") = DATETIME_STAMP

            rowICTIRTV1.Item("RTV_STATUS") = "H"
            rowICTIRTV1.Item("REGISTER_IND") = "0"
            dst.Tables("ICTIRTV1").Rows.Add(rowICTIRTV1)
        Else
            Fill_Record("ICTIRTV1", Absx1.txtFor("RTV_NO").Text)
            dst.AcceptChanges()

            If rowICTIRTV1.Item("REVERSED_BY_RTV_NO") & "" <> "" Or rowICTIRTV1.Item("REVERSES_RTV_NO") & "" <> "" Then
                'do not allow reversal
                UltraExplorerBar1.Groups("Screen Control").Items("Reverse RTV").Settings.Enabled = DefaultableBoolean.False
            End If

            With dst.Tables("ICTIRTV0").Rows
                .Add(New String() {"Entered By", rowICTIRTV1.Item("INIT_OPER")})
                .Add(New String() {"Entered On", Format(rowICTIRTV1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                .Add(New String() {"Voucher No", rowICTIRTV1.Item("VOUCHER_NO") & ""})
                .Add(New String() {"Order No", rowICTIRTV1.Item("ORDR_NO") & ""})
            End With
        End If

        Fill_Records("APTVEND1", Absx1.txtFor("VEND_CODE").Text)

        Fill_Records("ICTIRTV2", New Object() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("RTV_NO").Text})
        Fill_Records("ICTIRTV3", Absx1.txtFor("RTV_NO").Text)
        Fill_Records("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)


        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record(Optional ByVal isReversal As Boolean = False)

        Call BeginTrans()
        Dim requiresFinalization As Boolean = (rowICTIRTV1.Item("RTV_STATUS") = "F")

        If isReversal Then
            Set_Up_Reversal()
        End If

        rowICTIRTV1 = dst.Tables("ICTIRTV1").Rows(0)
        ICCMAIN1.Update_RTV(Me, rowICTIRTV1)

        If requiresFinalization Then
            ICCMAIN1.Finalize_RTV(rowICTIRTV1.Item("RTV_NO"), "", "")
        End If
        CommitTrans("Update Complete")
 
    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View"
                Absx1.txtFor("RTV_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        If Me.MENU_ITEM_OBJECT = "ICFIRTV1" Then
            Call Load_Popup_Menu(grdICTIRTVX, "SSB", "Show Filter", "Show GroupBox", "Finalize RTV(s)")
        Else
            Call Load_Popup_Menu(grdICTIRTVX, "SS", "Show Filter", "Show GroupBox")
        End If

    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            If grd.Name = "grdICTIRTVX" And Not InquiryMode Then
                Dim isFinalizeVisible As Boolean = True
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If row.Cells("RTV_STATUS").Value <> "H" Then
                        isFinalizeVisible = False
                    End If
                Next
                DirectCast(tlb_pop.Tools("Finalize RTV(s)"), UltraWinToolbars.ButtonTool).SharedProps.Visible = isFinalizeVisible
            End If

        End If
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick
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

            Case "Finalize RTV(s)"

                Dim finalizeCount As Integer = grd.Selected.Rows.Count

                If finalizeCount > 1 Then
                    If MessageBox.Show(String.Format("Are you sure you want to finalize multiple ({0}) RTVs?  Each RTV will receive the same RA number and tracking number.", finalizeCount), "Finalize Multiple RTVs", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Cancel Then
                        Exit Sub
                    End If
                End If
                Dim frmRtvInfo As New ICFIRTVF() 'grd.ActiveRow.Cells("RTV_NO").Value

                If frmRtvInfo.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    BeginTrans()
                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim allowUpdate As Boolean = (ASCDATA1.GetDataValue("SELECT RTV_STATUS FROM ICTIRTV1 WHERE RTV_NO=:PARM1 FOR UPDATE", "V", New String() {row.Cells("RTV_NO").Value}) <> "F")
                        If allowUpdate Then
                            row.Cells("VOUCHER_NO").Value = ICCMAIN1.Finalize_RTV(row.Cells("RTV_NO").Value, frmRtvInfo.RANo, frmRtvInfo.TrackingNo)
                            row.Cells("RTV_RA_NO").Value = frmRtvInfo.RANo
                            row.Cells("RTV_TRACKING_NO").Value = frmRtvInfo.TrackingNo
                        End If
                        row.Cells("RTV_STATUS").Value = "F"
                    Next
                    CommitTrans()
                    MsgBox("Finalization Complete")
                End If

                frmRtvInfo.Dispose()

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("WHSE_CODE").TextLength > 0 Then
                        Call Click_Command("New", e)
                    End If
                End If
            Case "RTV_NO"
                    If e.KeyCode = Windows.Forms.Keys.Enter Then
                        Call Click_Command("View", e)
                    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VEND_CODE"
                If Not InquiryMode And Absx1.txtFor("WHSE_CODE").TextLength > 0 Then
                    Call Click_Command("New")
                End If
            Case "RTV_NO"
                Call Click_Command("View")
        End Select
    End Sub

#End Region


#Region "grdICTIRTV2"

    Private Sub grdICTIRTV2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIRTV2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTIRTV2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = dst.Tables("ICTIRTV1").Rows(0).Item("WHSE_CODE")
                    Dim rowICTSTAT2 = Fill_Record("ICTSTAT2", New String() {ITEM_CODE, WHSE_CODE}, True)
                    Dim PRICE_CATGY_CODE As String = cdr.Item("PRICE_CATGY_CODE") & ""
                    Dim rowICTPCAT1 As DataRow = LookUp("ICTPCAT1", PRICE_CATGY_CODE)
                    Dim PROD_CATGY_CODE As String = rowICTPCAT1.Item("PROD_CATGY_CODE") & ""
                    Dim PO_COST As Decimal = Val(rowICTPCAT1.Item("PO_COST") & "")
                    Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", PRICE_CATGY_CODE, True)
                    Dim PRICE_CATGY_COST_TOTAL As Decimal = Val(rowICTCOSTC.Item("PRICE_CATGY_COST_TOTAL") & "")
                    e.Cell.Row.Cells("PRICE_CATGY_CODE").Value = PRICE_CATGY_CODE
                    e.Cell.Row.Cells("WHSE_QTY_ON_HAND").Value = Convert.ToInt32(ASCDATA1.GetDataValue("SELECT NVL(SUM(WHSE_QTY_ON_HAND),0) FROM ICTSTAT2 WHERE WHSE_CODE=:PARM1 AND ITEM_CODE=:PARM2", "VV", New Object() {WHSE_CODE, ITEM_CODE}))
                    e.Cell.Row.Cells("PO_COST").Value = PO_COST
                    e.Cell.Row.Cells("PROD_CATGY_CODE").Value = PROD_CATGY_CODE
                    e.Cell.Row.Cells("PRICE_CATGY_COST_TOTAL").Value = PRICE_CATGY_COST_TOTAL

                Else
                    grdICTIRTV2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

        End Select
    End Sub


    Private Sub grdICTIRTV2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIRTV2.AfterRowActivate
        With grdICTIRTV2.DisplayLayout.Bands(0)
            If grdICTIRTV2.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTIRTV2.ActiveCell = grdICTIRTV2.ActiveRow.Cells("ITEM_CODE")
                grdICTIRTV2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If

        End With

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdICTIRTV2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIRTV2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIRTV2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIRTV2.AfterRowUpdate
        DisplayTotals()
        If e.Row.IsAddRow Then
            grdICTIRTV2.Rows.TemplateAddRow.Cells("ITEM_CODE").Activate()
        End If
        'If e.Row.Cells("ITEM_ORDER_CODE").Text = "P" Then
        '    e.Row.Cells("ITEM_CODE").Appearance.BackColor = Color.Yellow
        'End If
    End Sub

    Private Sub grdICTIRTV2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTIRTV2.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                'cdr = LookUp("ICTITEM1", e.Cell.Text)
                'If cdr Is Nothing Then
                '    grdICTIRTV2.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
                '    e.Cancel = True
                'End If

        End Select

    End Sub

    Private Sub grdICTIRTV2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIRTV2.BeforeExitEditMode

        If grdICTIRTV2.ActiveRow Is Nothing Then
            Exit Sub
        End If

        With grdICTIRTV2.ActiveCell

            Select Case .Column.Key
                Case "ITEM_CODE"
                    If .Text <> "" Then
                        .Value = .Text.ToUpper
                    End If
                    If .Text <> "" Then
                        'cdr = LookUp("ICTITEM1", .Text)
                        cdr = Fill_Record("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Item Code (" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            .Value = cdr.Item("ITEM_CODE")
                            Dim rowICTPCAT1 As DataRow = LookUp("ICTPCAT1", cdr.Item("PRICE_CATGY_CODE") & "")
                            If rowICTPCAT1 Is Nothing Then
                                ASCMAIN1.Progress("Item " & .Text & " has an invalid Price Category (" & cdr.Item("PRICE_CATGY_CODE") & ")")
                                e.Cancel = True
                            Else
                                If rowICTPCAT1.Item("VEND_CODE") <> HFs("VEND_CODE") Then
                                    ASCMAIN1.Progress("Item " & .Text & " belongs to a different manufacturer (" & rowICTPCAT1.Item("VEND_CODE") & ")")
                                    e.Cancel = True
                                End If
                                If rowICTPCAT1.Item("PO_COST") = 0 And rowICTPCAT1.Item("PRICE_CATGY_SAMPLE_IND") = "0" Then
                                    ASCMAIN1.Progress("Item " & .Text & " is a non-sample with zero cost. See supervisor.")
                                End If
                            End If
                        End If
                    End If
            End Select
        End With

    End Sub

    Private Sub grdICTIRTV2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIRTV2.BeforeRowUpdate
        With grdICTIRTV2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                Call LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    e.Row.Cells("ITEM_ORDER_CODE").Value = cdr.Item("ITEM_ORDER_CODE")
                End If
            End If

            Dim rowICTIRTV2() As DataRow = dst.Tables("ICTIRTV2") _
                .Select("ITEM_CODE = '" & e.Row.Cells("ITEM_CODE").Text & "'")

            If rowICTIRTV2 IsNot Nothing AndAlso rowICTIRTV2.Length > 0 AndAlso e.Row.IsAddRow Then
                Dim AddQty As Integer = If(e.Row.Cells("RTV_QTY").Value Is DBNull.Value, 1, e.Row.Cells("RTV_QTY").Value)
                rowICTIRTV2(0).Item("RTV_QTY") = Val(rowICTIRTV2(0).Item("RTV_QTY") & "") + AddQty
                e.Cancel = True
            Else
                If e.Row.Cells("RTV_QTY").Value Is DBNull.Value Then
                    e.Row.Cells("RTV_QTY").Value = 1
                End If
                If Val(e.Row.Cells("RTV_QTY").Text) <= 0 Then
                    MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("RTV_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
                grdICTIRTV2.ActiveCell = grdICTIRTV2.Rows.TemplateAddRow.Cells("ITEM_CODE")
            End If

            If Not e.Cancel Then
                If e.Row.Cells("RTV_NO").Text = "" Then
                    .ActiveRow.Cells("RTV_NO").Value = Absx1.CtlFor("RTV_NO").Text
                    .ActiveRow.Cells("RTV_LNO").Value = Val(dst.Tables("ICTIRTV2").Compute("Max(RTV_LNO)", "") & "") + 1
                End If
            End If

            If e.Row.Cells("ITEM_ORDER_CODE").Text = "P" Then
                e.Row.Cells("ITEM_CODE").Appearance.BackColor = Color.Yellow
            End If
        End With
        DisplayTotals()
    End Sub

    Private Sub grdICTIRTV2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIRTV2.ClickCellButton

        If grdICTIRTV2.ActiveRow Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim sql_where As String = "EXISTS (SELECT 1 FROM ICTCATL1 WHERE ITEM_CODE=ICTITEM1.ITEM_CODE AND VEND_CODE = '" & HFs("VEND_CODE") & "')" 'vend_code

        Select Case grd.ActiveCell.Column.Key
            Case "ITEM_CODE"
                Call grdClickCellButton(grdICTIRTV2, sql_where, False)
                'grdICTIRTV2.ActiveCell.Row.Cells.Item("RTV_QTY").Value = 1
        End Select


    End Sub

    Private Sub grdICTIRTV2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIRTV2.Error
        If grdICTIRTV2.ActiveRow IsNot Nothing Then
            grdICTIRTV2.ActiveRow.CancelUpdate()
        End If
    End Sub

#End Region

    Sub DisplayTotals()
        Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIRTV2").Compute("SUM(LINE_COSTS)", "") & "")
        Dim PO_TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIRTV2").Compute("SUM(PO_LINE_COSTS)", "") & "")
        Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
        Absx1.numFor("PO_TOTAL_COSTS").Value = PO_TOTAL_COSTS
    End Sub

    Private Sub grdICTIRTVX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIRTVX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RTV_NO").Text = e.Row.Cells("RTV_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, _
  Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"

        End Select
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdICTIRTV3.DataSource = dst.Tables("ICTIRTV3")
                Dim dvw As DataView = dst.Tables("ICTIRTV3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIRTV3.DataSource = dst.Tables("ICTIRTV3")
                Dim dvw As DataView = dst.Tables("ICTIRTV3").DefaultView
                Dim RTV_LNO As Integer = 0
                If grdICTIRTV2.ActiveRow IsNot Nothing Then
                    RTV_LNO = Val(grdICTIRTV2.ActiveRow.Cells("RTV_LNO").Text)
                End If
                dvw.RowFilter = "RTV_LNO = " & CStr(RTV_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIRTV3").Clone
                Dim RTV_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIRTV3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("ICTIRTV3").Compute _
                    ("SUM(DIST_AMT)", _
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("RTV_NO") = Absx1.txtFor("RTV_NO").Text
                    row.Item("RTV_LNO") = 0
                    RTV_GNO += 1
                    row.Item("RTV_GNO") = RTV_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIRTV3.DataSource = tbl
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("ICTIRTVX", YP)
        If optRtvStatus.Value = "A" Then
            dst.Tables("ICTIRTVX").DefaultView.RowFilter = ""
        Else
            dst.Tables("ICTIRTVX").DefaultView.RowFilter = String.Format("RTV_STATUS = '{0}'", optRtvStatus.Value)
        End If
        grdICTIRTVX.Text = "Entered in " & cbeYP.Text
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String, _
                       ByVal WHSE_CODE As String, _
                       ByVal QTY_FIELD As String, _
                       ByVal S As Integer) As String

        Dim msg As String = ""

        Dim linecount As Integer = 0
        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from ICTSTAT2 where ITEM_CODE = '" & ITEM_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            Dim WHSE_QTY_ON_HAND As Integer = 0
            If rowICTSTAT2 IsNot Nothing Then
                WHSE_QTY_ON_HAND = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
            If WHSE_QTY_ON_HAND + S * QTY < 0 And linecount < 20 Then
                msg &= vbCr & Format("Item " & ITEM_CODE & " has only " & CStr(WHSE_QTY_ON_HAND) & " On Hand")
                linecount += 1
            ElseIf linecount >= 20 Then
                linecount += 1
            End If
        Next
        If linecount > 20 Then
            msg &= vbCr & "... and " & (linecount - 20).ToString() & " more items with insufficient quantity."
        End If

        Return msg
    End Function

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_RTV_NO = ASCMAIN1.Next_Control_No("ICTIRTV1.RTV_NO")

        Dim rowICTIRTV1_orig As DataRow = dst.Tables("ICTIRTV1").NewRow
        rowICTIRTV1_orig.ItemArray = rowICTIRTV1.ItemArray

        rowICTIRTV1 = dst.Tables("ICTIRTV1").Rows(0)
        rowICTIRTV1.Item("REVERSED_BY_RTV_NO") = REVERSED_BY_RTV_NO
        rowICTIRTV1.Item("RTV_STATUS") = "F"
        rowICTIRTV1.Item("RTV_DATE") = Today
        rowICTIRTV1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIRTV1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIRTV1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIRTV1")

        'dst.Tables("SOTRTRN1").Clear()
        'rowICTIRTV1 = dst.Tables("SOTRTRN1").NewRow
        rowICTIRTV1.ItemArray = rowICTIRTV1_orig.ItemArray
        'dst.Tables("SOTRTRN1").Rows.Add(rowICTIRTV1)
        rowICTIRTV1.AcceptChanges()
        rowICTIRTV1.SetAdded()

        'Set fields for reversal return
        With rowICTIRTV1
            .Item("REVERSES_RTV_NO") = .Item("RTV_NO")
            .Item("RTV_NO") = REVERSED_BY_RTV_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("TOTAL_COSTS") *= -1
            .Item("PO_TOTAL_COSTS") *= -1
            .Item("RTV_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
            .Item("JOURNAL_IND") = "0"
            .Item("JOURNAL_XNO") = DBNull.Value
            .Item("VOUCHER_NO") = DBNull.Value
            .Item("RTV_STATUS") = "F"
            '.Item("RTV_NOTE") = reversalReason
        End With

        'reversalReason = ""

        'Set new RTRN_NO and reverse all quantities for this return.
        For Each row As DataRow In dst.Tables("ICTIRTV2").Rows
            row.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        Next
        Update_Record_TDA("ICTIRTV2")

        For Each row As DataRow In dst.Tables("ICTIRTV2").Rows
            row.Item("RTV_NO") = REVERSED_BY_RTV_NO
            If row.Item("RTV_QTY") IsNot DBNull.Value Then
                row.Item("RTV_QTY") *= -1
            End If
            row.AcceptChanges()
            row.SetAdded()
        Next

        dst.Tables("ICTIRTV3").Clear()

    End Sub


    Sub Print_RTV()
        Print_Report_Begin()
        Generate_Report("ICRIRTVR", "Return to Vendor")
        Print_Report_End()
        Try
            clsASCBASE1.F.Activate()
            'F.Activate()
        Catch ex As Exception

        End Try
        'Print_Report_End(True, , ROWs("SOTPARM1").Item("SO_PARM_PTR_RTN_RPT") & "")
    End Sub

    Private Sub grdICTIRTV2_InitializeRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIRTV2.InitializeRow
        'If e.Row.Cells("ITEM_ORDER_CODE").Text = "P" Then
        '    e.Row.Cells("ITEM_CODE").Appearance.BackColor = Color.Yellow
        'End If
    End Sub

    Private Sub UltraOptionSet1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRtvStatus.ValueChanged
        If dst.Tables.Contains("ICTIRTVX") Then
            If optRtvStatus.Value = "A" Then
                dst.Tables("ICTIRTVX").DefaultView.RowFilter = ""
            Else
                dst.Tables("ICTIRTVX").DefaultView.RowFilter = String.Format("RTV_STATUS = '{0}'", optRtvStatus.Value)
            End If
        End If
    End Sub

    Private Sub grdICTIRTVX_InitializeRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIRTVX.InitializeRow
        If e.Row.Cells("REVERSED_BY_RTV_NO").Value & "" <> "" Or e.Row.Cells("REVERSES_RTV_NO").Value & "" <> "" Then
            e.Row.Appearance.ForeColor = Color.Red
        End If
    End Sub

    Private Sub grdICTIRTVX_AfterRowActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdICTIRTVX.AfterRowActivate
        If Not grdICTIRTVX.Selected.Rows.Contains(grdICTIRTVX.ActiveRow) Then
            If System.Windows.Forms.Control.ModifierKeys <> Keys.Shift And System.Windows.Forms.Control.ModifierKeys <> Keys.Control Then
                grdICTIRTVX.Selected.Rows.Clear()
                grdICTIRTVX.Selected.Rows.Add(grdICTIRTVX.ActiveRow)
            End If
        End If
    End Sub
End Class