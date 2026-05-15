Public Class APF1099U
    Dim APT1099U As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_APT1099U(True)

        With dst
            ASCMAIN1.sql = "Select * from " & APT1099U
            Create_TDA(.Tables.Add("APT1099U"), APT1099U, "**", 0, True, "", 1)

            Create_TDA(.Tables.Add, "APTVEND1", "*", 1, False)
        End With

        grdAPT1099U.DataSource = dst.Tables("APT1099U")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdAPT1099U.DisplayLayout.Bands(0).Columns
            If gcol.Key = "INV_1099_AMT" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        Create_Summary(grdAPT1099U, "VOUCHER_NO", "Count")
        Create_Summary(grdAPT1099U, New String() {"CHECK_AMT", "INV_AMT_APPLIED", "INV_1099_AMT", "INV_1099_AMT_ORIG"})

        Dim YYYYs As New List(Of String)
        YYYYs.Add(Now.Year)
        YYYYs.Add(CStr(Val(Now.Year) - 1))
        cbeYYYY.DataSource = YYYYs

        TABLE_NAME = "APTVEND1"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If cbeYYYY.Value & "" = "" Then
                    EMsg &= vbCr & "You Must First Specify a Valid Payment Year"
                End If

                If Absx1.txtFor("VEND_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Valid Vendor Code"
                Else
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If rowAPTVEND1 Is Nothing Then
                        EMsg &= vbCr & "You Must First Specify a Valid Vendor Code"
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

            Case "Load"
                EntryMode = "I"
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

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdAPT1099U.Visible = tf

        If ScreenMode Then
            'tabMain.SelectedTab = tabMain.Tabs("Vendor Name && Address")
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
        {"APT1099U"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        EnforceConstraints(False)
        Create_APT1099U(False)
        'Fill_Records("APT1099U", New String() _
        '                 {Absx1.txtFor("VEND_CODE").Text, _
        '                  Absx1.cbeFor("YYYY").Text})
        Fill_Records("APT1099U")
        Fill_Records("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
        EnforceConstraints(True)

        grdAPT1099U.Text = "Vouchers Appearing on Checks with Check Date recorded " & Absx1.cbeFor("YYYY").Value

    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("APT1099U")

        ASCMAIN1.sql = "" _
        & "Begin " _
        & " Declare Cursor C1 is Select * from " & APT1099U & ";" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "    Update APTINVH1 Set INV_1099_AMT = R1.INV_1099_AMT where VOUCHER_NO = R1.VOUCHER_NO; " _
        & "  End Loop; " _
        & " End; " _
        & "End;"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdAPT1099U, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "1099 All", "De-1099 All")

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

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "1099 All", "De-1099 All"
                For Each rowAPT1099U As DataRow In dst.Tables("APT1099U").Select
                    rowAPT1099U.Item("INV_1099_AMT") = IIf(e.Tool.Key = "1099 All", rowAPT1099U.Item("INV_AMT_APPLIED"), 0)
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Shipper Invoice Inquiry"
            '    Dim SHIPPER_CTL_NO As String = grd.ActiveRow.Cells("SHIPPER_CTL_NO").Value
            '    Context_Launch("Load", SHIPPER_CTL_NO, e.Tool.Key, "SOFSHIP1"

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Absx1.txtFor("VEND_CODE").Text = UCase(Absx1.txtFor("VEND_CODE").Text)
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VEND_CODE"
                Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                Absx1.txtFor("VEND_CODE").Text = UCase(Absx1.txtFor("VEND_CODE").Text)
        End Select

    End Sub

    Public Overrides Sub CheckedChanged_Special(ByVal COLUMN_NAME As String, ByVal chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor)
        MyBase.CheckedChanged_Special(COLUMN_NAME, chk)

        Select Case COLUMN_NAME
            'Case "JOINT_VENTURE"
            ' splJV.Panel2Collapsed = Not (chk.Checked)
        End Select

    End Sub
#End Region

    Overrides Sub Prepare_for_View_Lookup_Special( _
     ByVal ctl As Control, _
     ByVal COLUMN_NAME As String, _
     Optional ByRef sql_where As String = "", _
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            'Case "VEND_CODE"
            '    If MENU_ITEM_PP = "S" Then
            '        sql_where = "VEND_TYPE = 'M'"
            '    End If
        End Select
    End Sub


    Sub Create_APT1099U(ByVal initialize As Boolean)

        Dim YYYY As String = "0000"
        If Not initialize Then YYYY = Absx1.cbeFor("YYYY").Value

        Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text

        ASCMAIN1.sql = " Select" _
        & "  APTCHCK2.VOUCHER_NO" _
        & ", APTCHCK1.CHECK_DATE" _
        & ", APTCHCK2.CHECK_NUM" _
        & ", APTCHCK2.INV_NUM" _
        & ", APTCHCK2.INV_DATE" _
        & ", APTCHCK1.CHECK_AMT" _
        & ", APTCHCK2.INV_AMT_APPLIED" _
        & ", APTINVH1.INV_1099_AMT" _
        & ", APTINVH1.INV_1099_AMT INV_1099_AMT_ORIG" _
        & " from APTCHCK1, APTCHCK2, APTINVH1"


        If initialize Then
            ASCMAIN1.sql &= " where ROWNUM < 1"
            APT1099U = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APT1099U & " Add Primary Key (VOUCHER_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & APT1099U)
            Dim DTE1 As Date = CDate("01/01/" & Absx1.cbeFor("YYYY").Value)
            Dim DTE2 As Date = DTE1.AddYears(1).AddDays(-1)

            ASCMAIN1.sql &= "" _
            & " where APTCHCK1.VEND_CODE = :PARM1" _
            & "   and APTCHCK1.CHECK_DATE >= :PARM2" _
            & "   and APTCHCK1.CHECK_DATE <= :PARM3" _
            & "   and APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE " _
            & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM " _
            & "   and APTCHCK1.CHECK_STATUS = 'I'" _
            & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO"

            ASCDATA1.ExecuteSQL("Insert into " & APT1099U & " " & ASCMAIN1.sql, _
                                "VDD", _
                                New Object() {Absx1.txtFor("VEND_CODE").Text, _
                                               DTE1, DTE2})
        End If
    End Sub

    Private Sub grdAPT1099U_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPT1099U.ClickCellButton
        e.Cell.Row.Cells("INV_1099_AMT").Value = e.Cell.Row.Cells("INV_AMT_APPLIED").Value
        e.Cell.Row.Update()
    End Sub

    Private Sub grdAPT1099U_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPT1099U.InitializeRow
        e.Row.Cells("MOVE").Value = "->"
    End Sub
End Class