Imports System.Drawing
Imports System.Math

Public Class ICFIXFR1

    Dim rowICTIXFR1 As DataRow
    Dim location_support As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIXFRI" Then
            InquiryMode = True
        End If

        With dst
            ASCMAIN1.sql = "Select ICTIXFR1.*" _
            & " from ICTIXFR1 where ICTIXFR1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTIXFRX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTIXFR1", "*")

            ASCMAIN1.sql = "Select ICTIXFR2.*, ICTITEM1.ITEM_DESC" _
            & " from ICTIXFR2,ICTITEM1 where ICTITEM1.ITEM_CODE = ICTIXFR2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTIXFR2", "**", 1)
            .Tables("ICTIXFR2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(XFR_QTY,0) * ISNULL(ITEM_COST_STD,0)")

            ASCMAIN1.sql = "Select ICTIXFR3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIXFR3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIXFR3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIXFR3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("ICTIXFR0")
            .Tables("ICTIXFR0").Columns.Add("KEY")
            .Tables("ICTIXFR0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTCATG1"
            Create_TDA(.Tables.Add, "ICTCATG1", "**", 0, False)

        End With

        Set_Read_Only(grpTotals, True)

        Fill_Records("ICTCATG1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        grdICTIXFR0.DataSource = dst.Tables("ICTIXFR0")
        grdICTIXFR2.DataSource = dst.Tables("ICTIXFR2")
        grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
        grdICTIXFRX.DataSource = dst.Tables("ICTIXFRX")

        Create_Summary(grdICTIXFRX, "XFR_NO", "Count")
        Create_Summary(grdICTIXFRX, "TOTAL_COSTS")

        Create_Summary(grdICTIXFR2, "XFR_LNO", "Count")
        Create_Summary(grdICTIXFR2, "XFR_QTY")
        Create_Summary(grdICTIXFR2, "LINE_COSTS")

        Create_Summary(grdICTIXFR3, "XFR_GNO", "Count")
        Create_Summary(grdICTIXFR3, "DIST_AMT")


        With grdICTIXFRX.DisplayLayout.Bands("ICTIXFRX")
            .Columns("XFR_NO").Header.Fixed = True
        End With

        'ASCMAIN1.Add_Value_List(grdICTIXFRX, "WHSE_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 where REASON_TYPE = 'A' order by REASON_DESC")

        grdICTIXFR0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdICTIXFR3, "ICTIXFR3")

        Set_Read_Only(grpTotals, True)
        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "WS") = 0 Then
            grpTotals.Visible = False
            With grdICTIXFR2.DisplayLayout.Bands(0)
                '.Columns("PRICE_CATGY_COST_TOTAL").Hidden = True
                .Columns("LINE_COSTS").Hidden = True
                .Columns("COST_CATGY_CODE").Hidden = True
                .Columns("PROD_CODE").Hidden = True
            End With
        End If

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
                Validate_Code("WHSE_CODE")

                If Absx1.dteFor("XFR_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If
                ' MULTITASKING

            Case "View"
                If Absx1.txtFor("XFR_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIXFR1 = LookUp("ICTIXFR1", Absx1.txtFor("XFR_NO").Text)
                    If rowICTIXFR1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("XFR_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                'If Absx1.txtFor("WHSE_CODE_TO").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify a Transfer-To Warehouse"
                'End If

                Dim DT As Date = Absx1.dteFor("XFR_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                Validate_Code("WHSE_CODE_TO")

                If Absx1.txtFor("WHSE_CODE").Text = Absx1.txtFor("WHSE_CODE_TO").Text Then
                    EMsg &= vbCr & "Transfer-From and Transfer-To Warehouses must be different." & vbCr
                End If

                If grdICTIXFR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTIXFR2 As DataRow In dst.Tables("ICTIXFR2").Select("", "", DataViewRowState.CurrentRows)
                        If rowICTIXFR2.Item("COST_CATGY_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Cost Category Code for " & rowICTIXFR2.Item("ITEM_CODE") & ""
                        End If
                        If rowICTIXFR2.Item("PROD_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Product Code for " & rowICTIXFR2.Item("ITEM_CODE") & ""
                        End If

                        If Val(rowICTIXFR2.Item("XFR_QTY") & "") <= 0 Then
                            EMsg &= vbCr & "Positive Values Only (see " & rowICTIXFR2.Item("ITEM_CODE") & ")"
                        End If
                    Next
                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTIXFR2")

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIXFR2", Absx1.txtFor("WHSE_CODE").Text, "XFR_QTY", -1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK to Continue Anyway?", MsgBoxStyle.YesNo, "The following Items do not have Sufficent Qty for this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Reverse"
                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Entire Warehouse"
                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                ASCMAIN1.sql = "Select Count (*) ITEMS, Sum (WHSE_QTY_ON_HAND) QTY from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0 and WHSE_CODE = '" & WHSE_CODE & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                Dim ITEMS As Int32 = Val(row.Item("ITEMS") & "")
                Dim QTY As Int32 = Val(row.Item("QTY") & "")

                If MsgBox("There are " & CStr(ITEMS) & " Items with non-zero On Hand in Warehouse " & WHSE_CODE _
                          & vbCrLf & "There is a total qty of " & Format(QTY, "#,##0") & " units On Hand in Warehouse " & WHSE_CODE _
                          & vbCrLf & vbCrLf & "Warehouse " & WHSE_CODE & " will be Empty (ie, On Hand = 0) after you click Update" _
                          & vbCrLf & vbCrLf & "If there are any items with Negative Qty On Hand they will NOT be transferred." _
                          & vbCrLf & vbCrLf & "If there are any Open POs, Open Sales Orders, Pick Tickets etc." _
                          & vbCrLf & " which are open with reference to this warehouse," _
                          & vbCrLf & " they will have to be cancelled independently from this transaction" _
                          & vbCrLf & vbCrLf & "OK to Transfer Qty On Hand of Entire Warehouse?", MsgBoxStyle.YesNo, _
                          "Verification") = MsgBoxResult.No Then
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reverse"
                Set_Up_Reversal()
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Entire Warehouse"
                Entire_Warehouse()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Entire Warehouse").Visible = (EntryMode = "N")

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V" AndAlso Not InquiryMode) _
                        AndAlso rowICTIXFR1 IsNot Nothing _
                        AndAlso rowICTIXFR1.Item("REVERSED_BY_XFR_NO") Is DBNull.Value _
                        AndAlso rowICTIXFR1.Item("REVERSES_XFR_NO") Is DBNull.Value

                End With

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "WM") <> 0
                .Groups("Show if Entered in").Visible = Not ScreenMode And InStr(ASCMAIN1.USER_SECURITY_CODEs, "WM") <> 0
                .Groups("Transfer Price Catgy").Visible = ScreenMode AndAlso (EntryMode = "N") _
                    AndAlso InStr(ASCMAIN1.USER_SECURITY_CODEs, "WM") <> 0 AndAlso Not InquiryMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'SplitContainer1.Panel2Collapsed = Not ScreenMode
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode
        grdICTIXFRX.Visible = Not ScreenMode

        If ScreenMode Then
            grdICTIXFR0.Visible = (EntryMode = "V")
            SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "WM") = 0
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                With grdICTIXFR2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
                With grdICTIXFR2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("XFR_QTY").CellAppearance.BackColor = Color.LightYellow
                End With
                With grdICTIXFR3.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End With
            Else
                With grdICTIXFR2.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                End With
                With grdICTIXFR2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("XFR_QTY").CellAppearance.BackColor = Color.Empty
                End With
                With grdICTIXFR3.DisplayLayout.Override
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

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTIXFR0", "ICTIXFR1", "ICTIXFR2", "ICTIXFR3", "ICTSTAT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.dteFor("XFR_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("XFR_NO").Text = ""
        Absx1.txtFor("PRICE_CATGY_CODE").Text = ""

        optGL.Tag = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'If EntryMode = "N" Then
        '    Absx1.txtFor("XFR_NO").Text = ASCMAIN1.Next_Control_No("SOTINVH1.XFR_NO")
        'End If

        If EntryMode = "N" Then
            rowICTIXFR1 = dst.Tables("ICTIXFR1").NewRow
            rowICTIXFR1.Item("XFR_NO") = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
            rowICTIXFR1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTIXFR1.Item("XFR_DATE") = HFs("XFR_DATE")
            rowICTIXFR1.Item("XFR_SOURCE") = "E"
            rowICTIXFR1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIXFR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIXFR1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIXFR1.Item("REGISTER_IND") = "0"
            rowICTIXFR1.Item("JOURNAL_IND") = "0"
            dst.Tables("ICTIXFR1").Rows.Add(rowICTIXFR1)
        Else
            Fill_Record("ICTIXFR1", Absx1.txtFor("XFR_NO").Text)
            dst.AcceptChanges()

            dst.Tables("ICTIXFR0").Rows.Add(New String() {"Entered By", rowICTIXFR1.Item("INIT_OPER")})
            dst.Tables("ICTIXFR0").Rows.Add(New String() {"Entered On", Format(rowICTIXFR1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
            dst.Tables("ICTIXFR0").Rows.Add(New String() {"Source", rowICTIXFR1.Item("XFR_SOURCE")})

            If rowICTIXFR1.Item("REVERSED_BY_XFR_NO") & "" <> "" Then
                Dim row As DataRow = LookUp("ICTIXFR1", rowICTIXFR1.Item("REVERSED_BY_XFR_NO"))
                dst.Tables("ICTIXFR0").Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                dst.Tables("ICTIXFR0").Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                dst.Tables("ICTIXFR0").Rows.Add(New String() {"using", rowICTIXFR1.Item("REVERSED_BY_XFR_NO")})
            ElseIf rowICTIXFR1.Item("REVERSES_XFR_NO") & "" <> "" Then
                dst.Tables("ICTIXFR0").Rows.Add(New String() {"Reverses", rowICTIXFR1.Item("REVERSES_XFR_NO")})
            End If
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowICTIXFR1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTIXFR2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = True '  Not location_support
            .Columns("LOCATION_CODE").Hidden = Not location_support
        End With

        Fill_Records("ICTIXFR2", Absx1.txtFor("XFR_NO").Text)
        Fill_Records("ICTIXFR3", Absx1.txtFor("XFR_NO").Text)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        TAC.ICCMAIN1.Update_Transfer(Me)
        If location_support Then
            Update_WHTLOCBX("T")
        End If

        Dim rowICTWHSE1_WHSE_CODE_TO As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE_TO").Text)
        If rowICTWHSE1_WHSE_CODE_TO.Item("WHSE_LOCATOR") & "" = "1" Then
            Update_WHTLOCBX("X")
        End If

        CommitTrans("Update Complete")

    End Sub

    Sub Update_WHTLOCBX(TRAN_TYPE As String)
        Dim rowICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").Rows(0)
        TAC.ICCMAIN1.Update_WHTLOCBX(TRAN_TYPE, rowICTIXFR1.Item("XFR_NO"))
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
        Load_Popup_Menu(grdICTIXFRX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIXFR2, "B", "Item Status Inquiry")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

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
                'Case "Acknowledge w/Notes"
                '    Log_SetMode(True, True)
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "XFR_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not InquiryMode Then
                    Click_Command("New")
                End If
            Case "XFR_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTIXFR2"

    Private Sub grdICTIXFR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIXFR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTIXFR2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    Dim rowICTSTAT2 = Fill_Record("ICTSTAT2", New String() {ITEM_CODE, WHSE_CODE}, True)
                    Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    Dim rowICTCATG1 As DataRow = LookUp("ICTCATG1", COST_CATGY_CODE)


                    Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", ITEM_CODE, True)
                    Dim ITEM_COST_STD As Decimal = Val(rowICTCOSTC.Item("ITEM_COST_TOTAL") & "")

                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD

                Else
                    grdICTIXFR2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "XFR_QTY"
                'With grdICTIXFR2.ActiveRow
                '    .Cells("LINE_AMOUNT").Value = Val(.Cells("ORDR_QTY").Value & "") * Val(.Cells("ORDR_UNIT_PRICE").Value & "")
                'End With
        End Select
    End Sub

    Private Sub grdICTIXFR2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterExitEditMode
        Select Case grdICTIXFR2.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    Dim ACCT_CODE As String = grdICTIXFR2.ActiveCell.Text
            '    If ACCT_CODE <> "" Then
            '        grdICTIXFR2.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTJRNL2.ActiveCell.Column.Key)
            '    End If
        End Select
    End Sub

    Private Sub grdICTIXFR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterRowActivate
        With grdICTIXFR2.DisplayLayout.Bands(0)
            If grdICTIXFR2.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTIXFR2.ActiveCell = grdICTIXFR2.ActiveRow.Cells("ITEM_CODE")
                grdICTIXFR2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If

        End With

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdICTIXFR2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIXFR2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIXFR2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIXFR2.AfterRowUpdate
        DisplayTotals()
    End Sub

    'Private Sub grdICTIXFR2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTIXFR2.BeforeCellUpdate
    '    Select Case e.Cell.Column.Key
    '        Case "ITEM_CODE"
    '            'cdr = LookUp("ICTITEM1", e.Cell.Text)
    '            'If cdr Is Nothing Then
    '            '    grdICTIXFR2.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
    '            '    e.Cancel = True
    '            'End If

    '    End Select

    'End Sub

    Private Sub grdICTIXFR2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIXFR2.BeforeExitEditMode
        With grdICTIXFR2.ActiveCell
            Select Case .Column.Key
                Case "ITEM_CODE"
                    If .Text <> "" Then
                        .Value = .Text.ToUpper
                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Item Code (" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If

                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdICTIXFR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIXFR2.BeforeRowUpdate
        With grdICTIXFR2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If


            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                If e.Row.Cells("LOCATION_CODE").Text = "" Then
                    e.Cancel = True
                Else
                    LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")", _
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If

            End If

            If Val(e.Row.Cells("XFR_QTY").Text) <= 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("XFR_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("XFR_NO").Text = "" Then
                    .ActiveRow.Cells("XFR_NO").Value = Absx1.CtlFor("XFR_NO").Text
                    .ActiveRow.Cells("XFR_LNO").Value = Val(dst.Tables("ICTIXFR2").Compute("Max(XFR_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTIXFR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIXFR2.ClickCellButton

        If grdICTIXFR2.ActiveRow Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim sql_where As String = ""
        Select Case grd.ActiveCell.Column.Key
            Case "ITEM_CODE"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTIXFR2, sql_where, sql_where <> "")

    End Sub

    Private Sub grdICTIXFR2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIXFR2.Error
        grdICTIXFR2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIXFR2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTIXFRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIXFRX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("XFR_NO").Text = e.Row.Cells("XFR_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
                Dim dvw As DataView = dst.Tables("ICTIXFR3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIXFR3.DataSource = dst.Tables("ICTIXFR3")
                Dim dvw As DataView = dst.Tables("ICTIXFR3").DefaultView
                Dim XFR_LNO As Integer = 0
                If grdICTIXFR2.ActiveRow IsNot Nothing Then
                    XFR_LNO = Val(grdICTIXFR2.ActiveRow.Cells("XFR_LNO").Text)
                End If
                dvw.RowFilter = "XFR_LNO = " & CStr(XFR_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIXFR3").Clone
                Dim XFR_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIXFR3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("ICTIXFR3").Compute _
                    ("SUM(DIST_AMT)", _
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("XFR_NO") = Absx1.txtFor("XFR_NO").Text
                    row.Item("XFR_LNO") = 0
                    XFR_GNO += 1
                    row.Item("XFR_GNO") = XFR_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIXFR3.DataSource = tbl
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
        Fill_Records("ICTIXFRX", YP)
        Sort_grdColumns(grdICTIXFRX, "XFR_NO".ToLower)
        grdICTIXFRX.Text = "Entered in " & cbeYP.Text
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String, _
                       ByVal WHSE_CODE As String, _
                       ByVal QTY_FIELD As String, _
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from ICTSTAT2 where ITEM_CODE = '" & ITEM_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            Dim WHSE_QTY_ON_HAND As Integer = 0
            If rowICTSTAT2 IsNot Nothing Then
                WHSE_QTY_ON_HAND = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
            If WHSE_QTY_ON_HAND + S * QTY < 0 Then
                msg &= vbCr & Format("Item " & ITEM_CODE & " has only " & CStr(WHSE_QTY_ON_HAND) & " On Hand")
            End If

        Next

        Return msg
    End Function

    Private Sub btnLoadPCAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadPCAT.Click

        Dim WHSE_CODE As String = MyBase.Absx1.txtFor("WHSE_CODE").Text.Trim
        Dim WHSE_CODE_TO As String = MyBase.Absx1.txtFor("WHSE_CODE_TO").Text.Trim
        Dim PRICE_CATGY_CODE As String = MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Text.Trim
        Dim sql As String = String.Empty
        Dim ITEM_CODE As String = String.Empty

        If WHSE_CODE.Length = 0 Or WHSE_CODE_TO.Length = 0 Or PRICE_CATGY_CODE.Length = 0 Then
            MessageBox.Show("You must provide the Warehouse, Transfer To warehouse and Price Category Code values.", "Transfer", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Clear()
            Exit Sub
        End If

        Dim msg As String = "Do you want to transfer inventory from Warehouse: " & WHSE_CODE
        msg &= " to Warehouse: " & WHSE_CODE_TO
        msg &= " for Price Category Code: " & PRICE_CATGY_CODE & "?"

        If MessageBox.Show(msg, "Transfer", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Clear()
            Exit Sub
        End If

        Dim replaceQty As Boolean = False

        If dst.Tables("ICTIXFR2").Select("PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'").Length > 0 Then
            msg = "Price Category Code: " & PRICE_CATGY_CODE & " already exists in the items to be transferred. Do you want to overwrite the data?"
            msg &= Environment.NewLine & Environment.NewLine & "Click 'No' to abort the transfer."
            msg &= Environment.NewLine & Environment.NewLine & "Click 'Yes' to replace the transfer quantity with the quantity found in warehouse " & WHSE_CODE & "."

            Select Case MessageBox.Show(msg, "Transfer", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

                Case Windows.Forms.DialogResult.No
                    MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Clear()
                    Exit Sub

                Case Windows.Forms.DialogResult.Yes
                    replaceQty = True

            End Select
        End If

        sql = " SELECT ICTSTAT2.*"
        sql = sql & " FROM ICTSTAT2, ICTITEM1"
        sql = sql & " WHERE ICTITEM1.ITEM_CODE = ICTSTAT2.ITEM_CODE"
        sql = sql & " AND ICTITEM1.PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'"
        sql = sql & " AND ICTSTAT2.WHSE_CODE = '" & WHSE_CODE & "'"
        sql = sql & " AND WHSE_QTY_ON_HAND <> 0"

        For Each rowICTSTAT2 As DataRow In ASCDATA1.GetDataTable(sql).Rows
            ITEM_CODE = rowICTSTAT2("ITEM_CODE")

            If replaceQty AndAlso dst.Tables("ICTIXFR2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length > 0 Then
                While dst.Tables("ICTIXFR2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length > 0
                    dst.Tables("ICTIXFR2").Select("ITEM_CODE = '" & ITEM_CODE & "'")(0).Delete()
                End While
            End If

            grdICTIXFR2.DisplayLayout.Bands(0).AddNew.Activate()
            grdICTIXFR2.ActiveRow.Cells("ITEM_CODE").Value = ITEM_CODE
            grdICTIXFR2.ActiveRow.Cells("XFR_QTY").Value = rowICTSTAT2("WHSE_QTY_ON_HAND")
            grdICTIXFR2.ActiveRow.Update()
        Next

        MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Clear()

    End Sub

    Sub Entire_Warehouse()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_QTY_ON_HAND from ICTSTAT2 where WHSE_QTY_ON_HAND <> 0 and WHSE_CODE = '" & WHSE_CODE & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "ITEM_CODE")

            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim WHSE_QTY_ON_HAND As Int64 = row.Item("WHSE_QTY_ON_HAND")

            With grdICTIXFR2
                .Visible = False
                If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                    .ActiveRow = Nothing
                End If
                .DisplayLayout.Bands(0).AddNew()
                With .ActiveRow
                    .Cells("ITEM_CODE").Value = ITEM_CODE
                    '  .Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE
                    .Cells("XFR_QTY").Value = WHSE_QTY_ON_HAND
                    If location_support Then
                        .Cells("LOCATION_CODE").Value = rowICTWHSE1.Item("WHSE_LOC_SHP")
                    End If
                    .Update()
                End With
                .Visible = True
            End With
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Sort_grdColumns(grdICTIXFR2, "XFR_LNO")

    End Sub

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_XFR_NO = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
        Dim rowICTIXFR1_orig As DataRow = dst.Tables("ICTIXFR1").NewRow
        rowICTIXFR1_orig.ItemArray = rowICTIXFR1.ItemArray

        rowICTIXFR1 = dst.Tables("ICTIXFR1").Rows(0)
        rowICTIXFR1.Item("REVERSED_BY_XFR_NO") = REVERSED_BY_XFR_NO
        rowICTIXFR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIXFR1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIXFR1")

        rowICTIXFR1.ItemArray = rowICTIXFR1_orig.ItemArray
        rowICTIXFR1.AcceptChanges()
        rowICTIXFR1.SetAdded()

        With rowICTIXFR1
            .Item("REVERSES_XFR_NO") = .Item("XFR_NO")
            .Item("XFR_NO") = REVERSED_BY_XFR_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("XFR_DATE") = DATETIME_STAMP.Date
            .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
            .Item("JOURNAL_IND") = "0"
            .Item("JOURNAL_XNO") = DBNull.Value
        End With

        For Each row As DataRow In dst.Tables("ICTIXFR2").Rows
            row.Item("XFR_NO") = REVERSED_BY_XFR_NO
            If row.Item("XFR_QTY") IsNot DBNull.Value Then
                row.Item("XFR_QTY") *= -1
            End If
            If row.Item("OPS_YYYYPP") IsNot DBNull.Value Then
                row.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End If

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub
End Class