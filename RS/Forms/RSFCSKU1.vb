Public Class RSFCSKU1

    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim sqlRSTCSKUX As String
    Dim RSTCSKUW As String
    Dim sqlRSTCSKUW As String
    Dim RSTCSKUZ As String
    Dim RSTCSKUC As String

    Dim PIVOT_CODEs As New Dictionary(Of String, String)

    Dim RYW As String = ""
    Dim LYW As String = ""
    Dim RYP As String = ASCMAIN1.CYP
    Dim LYP As String = ASCMAIN1.CYP

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            If MENU_ITEM_OBJECT = "RSFCSKU1" Then

                ASCMAIN1.sql = "Select RSTCSKU1.*" & vbCrLf _
                     & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE" & vbCrLf _
                     & ", ICTCOLL1.HC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
                     & " from RSTCSKU1,ICTITEM1,ICTSIZEN,ICTCOLL1" & vbCrLf _
                     & " where ICTITEM1.ITEM_CODE = RSTCSKU1.ITEM_CODE" & vbCrLf _
                     & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
                     & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                     & "   and RSTCSKU1.CUST_CODE = :PARM1"
                Create_TDA(.Tables.Add, "RSTCSKU1", "**", 0, True, "V", 2)
                .Tables("RSTCSKU1").Columns.Add("SEL")
                .Tables("RSTCSKU1").Columns("SEL").DefaultValue = "0"

                .Tables("RSTCSKU1").Columns.Add("YW_MIN")
                .Tables("RSTCSKU1").Columns.Add("YW_MAX")
                .Tables("RSTCSKU1").Columns.Add("QTY_ONH", GetType(System.Int64))

                ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf _
                    & ", MIN (OPS_YYYYWW) YW_MIN, MAX(OPS_YYYYWW) YW_MAX" & vbCrLf _
                    & ", SUM (DECODE(OPS_YYYYWW,:PARM1,QTY_EOW,0)) QTY_ONH" & vbCrLf _
                    & " from RSTRETL1 where CUST_CODE = :PARM2 group by ITEM_CODE"
                Create_TDA(.Tables.Add, "RSTRETLX", "**", 0, False, "VV", 1)


            ElseIf MENU_ITEM_OBJECT = "RSFCSKUX" Then

                ASCMAIN1.sql = "Select * from ARTCUST2"
                Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, , , , "CUST_STORE_CLASS_CODE")

                sqlRSTCSKUX = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.MALL_CODE, ARTCUST2.CUST_STORE_CLASS_CODE" _
                    & " from ARTCUST1, ARTCUST2" _
                    & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE"
                ASCMAIN1.sql = sqlRSTCSKUX & " and ARTCUST2.CUST_CODE = :PARM1"
                Create_TDA(.Tables.Add, "RSTCSKUX", "**", 0, False, "V", 2)

            Else

                Get_PIVOT_CODEs()

                Create_WorkFile(True)

                ASCMAIN1.sql = "" _
                    & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
                    & ", SUM (SELLTHRU) SELLTHRU, SUM (ONHAND) ONHAND, SUM (SELLIN) SELLIN from (" & vbCrLf _
                    & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
                    & ", SUM (QTY_SOLD) SELLTHRU" & vbCrLf _
                    & ", SUM (DECODE(OPS_YYYYWW,:PARM1,QTY_EOW,0)) ONHAND" & vbCrLf _
                    & ", 0 SELLIN" & vbCrLf _
                    & " from RSTRETL1" & vbCrLf _
                    & " where OPS_YYYYPP = :PARM2" & vbCrLf _
                    & " and CUST_CODE = :PARM3" & vbCrLf _
                    & "   and ITEM_CODE in (Select ITEM_CODE from " & RSTCSKUW & " RSTCSKUW)" & vbCrLf _
                    & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
                    & " union " & vbCrLf _
                    & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
                    & ", 0 SELLTHRU" & vbCrLf _
                    & ", 0 ONHAND" & vbCrLf _
                    & ", SUM (ORDR_QTY_SHIP) SELLIN" & vbCrLf _
                    & " from SOTINVH2" & vbCrLf _
                    & " where ORDR_YYYYPP_UPDATED = :PARM4" & vbCrLf _
                    & "   and CUST_CODE = :PARM5" & vbCrLf _
                    & "   and ITEM_CODE in (Select ITEM_CODE from " & RSTCSKUW & " RSTCSKUW)" & vbCrLf _
                    & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
                    & ") group by CUST_CODE, CUST_STORE_NO, ITEM_CODE"
                Create_TDA(.Tables.Add, "RSTCSKUS", "**", 0, False, "VVVVV", 3)

                ASCMAIN1.sql = "Select RSTCSKUW.*, DECODE(RSTCSKU1.ITEM_CODE,NULL,'0','1') COMP" & vbCrLf _
                    & " from " & RSTCSKUW & " RSTCSKUW, RSTCSKU1" & vbCrLf _
                    & " where RSTCSKU1.CUST_CODE (+) = :PARM1 and RSTCSKU1.ITEM_CODE (+) = RSTCSKUW.ITEM_CODE"
                Create_TDA(.Tables.Add, "RSTCSKUW", "**", 0, False, "V", 1)

                ASCMAIN1.sql = "Select RSTCSKUZ.*, DECODE(RSTCSKU1.ITEM_CODE,NULL,'0','1') COMP" & vbCrLf _
                    & " from " & RSTCSKUZ & " RSTCSKUZ, RSTCSKU1" & vbCrLf _
                    & " where RSTCSKU1.CUST_CODE (+) = :PARM1 and RSTCSKU1.ITEM_CODE (+) = RSTCSKUZ.ITEM_CODE"
                Create_TDA(.Tables.Add, "RSTCSKUZ", "**", 0, False, "V", 2)

                ASCMAIN1.sql = "Select RSTCSKUC.* from " & RSTCSKUC & " RSTCSKUC"
                Create_TDA(.Tables.Add, "RSTCSKUC", "**", 0, False, "", 2)

                If MENU_ITEM_OBJECT = "RSFCSKUM" Then
                    'Create_Relation("RSTCSKUZ", "RSTCSKUW", "ITEM_CODE")
                    'Create_Relation("RSTCSKUW", "RSTCSKUZ", "ITEM_CODE")
                    Create_Relation("RSTCSKUC", "RSTCSKUZ", "CUST_STORE_CLASS_CODE")
                End If

                For Each TABLE_NAME As String In New String() {"RSTCSKUW", "RSTCSKUZ", "RSTCSKUC"}
                    With dst.Tables(TABLE_NAME).Columns
                        Dim iPIVOT_CODE As Integer = -1
                        For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                            iPIVOT_CODE += 1
                            Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                            For Each COLUMN_NAME As String In New String() {"SLS_QTY", "SLS_AMT"}
                                Dim CVAR As String = COLUMN_NAME & "_VAR" & sfx
                                Dim CVARPCT As String = COLUMN_NAME & "_VARPCT" & sfx
                                Dim CTY As String = COLUMN_NAME & "_TY" & sfx
                                Dim CLY As String = COLUMN_NAME & "_LY" & sfx
                                .Add(CVAR, GetType(System.Decimal), CTY & "-" & CLY)
                                .Add(CVARPCT, GetType(System.Decimal), "IIF(" & CLY & "=0,0," & CVAR & "/" & CLY & ")")
                                Dim CRANK As String = COLUMN_NAME & "_RANK" & sfx
                                .Add(CRANK, GetType(System.Int64))
                            Next
                        Next
                        .Add("OH", GetType(System.Int64), "OHST+OHDC")
                        .Add("WOS", GetType(System.Decimal), "IIF(SLS_L26=0 OR WKS=0,0,OH/(SLS_L26/WKS))")
                        .Add("WOS_ST", GetType(System.Decimal), "IIF(SLS_L26=0 OR WKS=0,0,OHST/(SLS_L26/WKS))")
                        .Add("WOS_DC", GetType(System.Decimal), "IIF(SLS_L26=0 OR WKS=0,0,OHDC/(SLS_L26/WKS))")
                    End With
                Next

            End If

        End With

        If MENU_ITEM_OBJECT = "RSFCSKU1" Then
            grdRSTCSKU1.DataSource = dst.Tables("RSTCSKU1")
            Format_grdRSTCSKU1()
        ElseIf MENU_ITEM_OBJECT = "RSFCSKUX" Then
            grdRSTCSKUX.DataSource = dst.Tables("RSTCSKUX")
            Format_grdRSTCSKUX()
        ElseIf MENU_ITEM_OBJECT = "RSFCSKUM" Then
            grdRSTCSKUW.DataSource = dst.Tables("RSTCSKUC")
            Format_grdRSTCSKUW()
        Else
            grdRSTCSKUW.DataSource = dst.Tables("RSTCSKUW")
            Format_grdRSTCSKUW()
        End If

        Set_cmbYW("RYW", ASCMAIN1.CYW, -5 * 52, 1 * 52, -1)

        btnExportByStore.Visible = (MENU_ITEM_OBJECT = "RSFCSKUM")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Select a Customer Code"
                Else
                    Validate_Code("CUST_CODE")
                    If EMsg = "" Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("RSTCSKU1", "CUST_CODE:" & CUST_CODE) Then Exit Sub
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
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
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not (EntryMode = "E")
                    .Items("Edit").Visible = (Not InquiryMode) And Not (EntryMode = "E") And (MENU_ITEM_OBJECT = "RSFCSKU1")
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V") And (MENU_ITEM_OBJECT = "RSFCSKU1")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V") And (MENU_ITEM_OBJECT = "RSFCSKU1")

                End With

                'Set_Read_Only_for_ctl(cbeYW, ScreenMode)
                .Groups("Filters").Visible = ScreenMode And (EntryMode = "V") And (MENU_ITEM_OBJECT <> "RSFCSKU1")
                .Groups("Comp SKUs").Visible = ScreenMode And (EntryMode = "E")

            End With
        End If

        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                grdRSTCSKUX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdRSTCSKUX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTCSKUX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                grdRSTCSKU1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTCSKU1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdRSTCSKUX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdRSTCSKUX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTCSKUX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

                grdRSTCSKU1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdRSTCSKU1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
        End If

        grdRSTCSKUX.Visible = ScreenMode And False ' (EntryMode <> "E")
        grdRSTCSKU1.Visible = ScreenMode And (MENU_ITEM_OBJECT = "RSFCSKU1")
        grdRSTCSKUW.Visible = ScreenMode And (EntryMode <> "E") And (MENU_ITEM_OBJECT <> "RSFCSKU1")

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpOptions, ScreenMode)
        ' spl.Panel1Collapsed = ScreenMode
        spl.SplitterDistance = 40

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        If MENU_ITEM_OBJECT = "RSFCSKU1" Then
            For Each TABLE_NAME As String In New String() {"RSTCSKU1", "RSTRETLX"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        ElseIf MENU_ITEM_OBJECT = "RSFCSKUX" Then
            For Each TABLE_NAME As String In New String() {"RSTCSKUX", "ARTCUST2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        Else
            For Each TABLE_NAME As String In New String() {"RSTCSKUW", "RSTCSKUZ", "RSTCSKUC"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End If

        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        RYW = cbeYW.Value
        LYW = ASCMAIN1.Week_Calc(RYW, -52)
        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        RYP = rowGLTPARM3.Item("YYYYPP")
        LYP = ASCMAIN1.Period_Calc(RYP, -12)

        Get_PIVOT_CODEs()

        EnforceConstraints(False)

        If MENU_ITEM_OBJECT = "RSFCSKU1" Then

            Fill_Records("RSTCSKU1", CUST_CODE)
            For Each rowRSTCSKU1 As DataRow In dst.Tables("RSTCSKU1").Select("")
                Get_1st_Last_Sale_YW_OH(rowRSTCSKU1)
                rowRSTCSKU1.Item("SEL") = "1"
            Next
            dst.Tables("RSTCSKU1").AcceptChanges()

        ElseIf MENU_ITEM_OBJECT = "RSFCSKUX" Then
            Fill_Records("RSTRETLX", New String() {RYW, CUST_CODE})

            Fill_Records("RSTCSKUX", New String() {CUST_CODE})
            grdRSTCSKUX.Text = "Store Classes for " & CUST_CODE
            Sort_grdColumns(grdRSTCSKUX, "CUST_CODE,CUST_STORE_NO")
            dst.Tables("RSTCSKUX").AcceptChanges()

            '   Fill_Records("ARTCUST2", CUST_CODE)
        Else

            Create_WorkFile(False)

            Fill_Records("RSTCSKUW", CUST_CODE)
            Fill_Records("RSTCSKUS", New String() {RYW, RYP, CUST_CODE, RYP, CUST_CODE})

            Fill_Records("RSTCSKUZ", CUST_CODE)
            Fill_Records("RSTCSKUC")

            Rank_SKUs()
            Set_COMP()

            If MENU_ITEM_OBJECT = "RSFCSKUM" Then
                Sort_grdColumns(grdRSTCSKUW, "CUST_STORE_CLASS_CODE")
                Sort_grdColumns(grdRSTCSKUW, "ITEM_CODE", , 1)
            Else
                Sort_grdColumns(grdRSTCSKUW, "ITEM_CODE")
            End If

            Select Case MENU_ITEM_OBJECT
                Case "RSFCSKUW"
                    grdRSTCSKUW.Text = "Weekly Sales by Store Class as of " & cbeYW.Text
                    grdRSTCSKUW.DisplayLayout.Bands(0).Groups("Inventory").Header.Caption = "Inventory @" & cbeYW.Text

                    Dim iPIVOT_CODE As Integer = -1
                    For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                        iPIVOT_CODE += 1
                        Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                        grdRSTCSKUW.DisplayLayout.Bands(0).Columns("SLS_QTY_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                        grdRSTCSKUW.DisplayLayout.Bands(0).Columns("SLS_AMT_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                    Next

                Case "RSFCSKUM"
                    grdRSTCSKUW.Text = "Rolling 12-Month Sales as of " & cbeYW.Text
                    Dim iPIVOT_CODE As Integer = -1
                    For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                        iPIVOT_CODE += 1
                        Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                        grdRSTCSKUW.DisplayLayout.Bands(0).Groups(Format(iPIVOT_CODE, "000")).Header.Caption = PIVOT_CODEs(PIVOT_CODE)
                        grdRSTCSKUW.DisplayLayout.Bands(1).Groups(Format(iPIVOT_CODE, "000")).Header.Caption = PIVOT_CODEs(PIVOT_CODE)

                        grdRSTCSKUW.DisplayLayout.Bands(0).Columns("SLS_QTY_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                        grdRSTCSKUW.DisplayLayout.Bands(0).Columns("SLS_AMT_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                        grdRSTCSKUW.DisplayLayout.Bands(1).Columns("SLS_QTY_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                        grdRSTCSKUW.DisplayLayout.Bands(1).Columns("SLS_AMT_LY" & sfx).Header.Caption = IIf(chkUSEPLAN.Checked, "Plan", "#LY")
                    Next
                    grdRSTCSKUW.DisplayLayout.Bands(0).Groups(1).Header.Caption = "YTD"
                    grdRSTCSKUW.DisplayLayout.Bands(1).Groups(1).Header.Caption = "YTD"

                    grdRSTCSKUW.DisplayLayout.Bands(0).Groups("Inventory").Header.Caption = "Inventory @" & cbeYW.Text
                    grdRSTCSKUW.DisplayLayout.Bands(1).Groups("Inventory").Header.Caption = "Inventory @" & cbeYW.Text
            End Select

        End If

        EnforceConstraints(True)
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Comp SKUs")

        ASCDATA1.DeleteRows("RSTCSKU1", "SEL<>'1'")

        Try
            BeginTrans()
            Update_Record_TDA("RSTCSKU1")
            'Update_Record_TDA("ARTCUST2")

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTCSKUX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTCSKU1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTCSKUW, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then ' Or Not GRDs.ContainsKey(e.SourceControl.Name) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing ' GRDs(Mid(e.SourceControl.Name, 4))
        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            e.Cancel = True
            Exit Sub
        Else
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdRSTCSKUX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim rowsSkipped As Int16 = 0

        Select Case e.Tool.Key

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If txtctl.Text <> "" Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                'If EntryMode = "" Then
                '    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                '        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                '        If cdr IsNot Nothing Then

                '        End If
                '    End If
                'End If
        End Select
    End Sub

#End Region

    Sub Create_WorkFile(initialize As Boolean)

        Dim RYP_01 As String = ASCMAIN1.Period_Calc(RYP, -11)
        Dim LYP_01 As String = ASCMAIN1.Period_Calc(LYP, -11)

        Dim TY_YP As String = ""
        Dim LY_YP As String = ""
        Dim sqlD As String = ""
        Dim iPIVOT_CODE As Integer = -1
        For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
            'Dim sfx As String = IIf(CUST_STORE_CLASS_CODE = "", "", "_" & CUST_STORE_CLASS_CODE)
            iPIVOT_CODE += 1
            Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
            Select Case MENU_ITEM_OBJECT
                Case "RSFCSKUW"
                    sqlD &= "" _
                        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_QTY_TY" & sfx & vbCrLf _
                        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & LYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_QTY_LY" & sfx & vbCrLf _
                        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTRETL1.AMT_SOLD ELSE 0 END) SLS_AMT_TY" & sfx & vbCrLf _
                        & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & LYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTRETL1.AMT_SOLD ELSE 0 END) SLS_AMT_LY" & sfx & vbCrLf

                Case "RSFCSKUM"
                    Dim X As String = "CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '000001' AND '000002' THEN RSTRETL1.XXX ELSE 0 END"
                    If iPIVOT_CODE = 0 Then
                        X = Replace(X, "000001", "000000")
                        TY_YP = RYP
                        LY_YP = LYP
                    Else
                        TY_YP = ASCMAIN1.Period_Calc(RYP, -12 + iPIVOT_CODE)
                        LY_YP = ASCMAIN1.Period_Calc(LYP, -12 + iPIVOT_CODE)
                    End If

                    sqlD &= "" _
                        & ", SUM (" & Replace(Replace(Replace(Replace(X, "000000", RYP_01), "000001", TY_YP), "000002", TY_YP), "XXX", "QTY_SOLD") & ") SLS_QTY_TY" & sfx & vbCrLf _
                        & ", SUM (" & Replace(Replace(Replace(Replace(X, "000000", LYP_01), "000001", LY_YP), "000002", LY_YP), "XXX", "QTY_SOLD") & ") SLS_QTY_LY" & sfx & vbCrLf _
                        & ", SUM (" & Replace(Replace(Replace(Replace(X, "000000", RYP_01), "000001", TY_YP), "000002", TY_YP), "XXX", "AMT_SOLD") & ") SLS_AMT_TY" & sfx & vbCrLf _
                        & ", SUM (" & Replace(Replace(Replace(Replace(X, "000000", LYP_01), "000001", LY_YP), "000002", LY_YP), "XXX", "AMT_SOLD") & ") SLS_AMT_LY" & sfx & vbCrLf

                    ', SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '201504' AND '201603' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_QTY_TY
                    ', SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '201404' AND '201503' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_QTY_LY
                    ', SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '201504' AND '201603' THEN RSTRETL1.AMT_SOLD ELSE 0 END) SLS_AMT_TY
                    ', SUM (CASE WHEN RSTRETL1.OPS_YYYYPP BETWEEN '201404' AND '201503' THEN RSTRETL1.AMT_SOLD ELSE 0 END) SLS_AMT_LY

                    ' DOES THE MONTH RANGE CLIP CORNERS ON THE WEEKS?
                    If iPIVOT_CODE = 0 Then
                        sqlD = Replace(sqlD, RYP_01, Mid(TY_YP, 1, 4) & "01")
                        sqlD = Replace(sqlD, LYP_01, Mid(LY_YP, 1, 4) & "01")
                        'sqlD = Replace(sqlD, TY_YP & "' THEN ", TY_YP & "' and OPS_YYYYWW >= '" & Mid(RYW, 1, 4) & "01" & "' and OPS_YYYYWW <= '" & RYW & "' THEN ")
                        'sqlD = Replace(sqlD, LY_YP & "' THEN ", LY_YP & "' and OPS_YYYYWW >= '" & Mid(LYW, 1, 4) & "01" & "' and OPS_YYYYWW <= '" & LYW & "' THEN ")
                    End If
            End Select
        Next

        Dim sqlRange As String = ""
        Select Case MENU_ITEM_OBJECT
            Case "RSFCSKUW"
                sqlRange = " and RSTRETL1.OPS_YYYYWW BETWEEN '" & Mid(LYW, 1, 4) & "01" & "' and '" & RYW & "'"
            Case "RSFCSKUM"
                sqlRange = " and RSTRETL1.OPS_YYYYPP BETWEEN '" & LYP_01 & "' and '" & RYP & "'"
        End Select

        Dim sqlX As String = "" _
            & "Select RSTRETL1.ITEM_CODE, MIN (RSTRETL1.OPS_YYYYWW) MINYW, MAX (RSTRETL1.OPS_YYYYWW) MAXYW" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' AND NVL(ARTCUST2.CUST_DC_IND,'0') = '1' THEN RSTRETL1.QTY_EOW ELSE 0 END) OHDC" & vbCrLf _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' AND NVL(ARTCUST2.CUST_DC_IND,'0') = '0' THEN RSTRETL1.QTY_EOW ELSE 0 END) OHST" & vbCrLf _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW BETWEEN '" & ASCMAIN1.Week_Calc(RYW, -26 + 1) & "' AND '" & RYW & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_L26" & vbCrLf _
            & ", MIN (CASE WHEN RSTRETL1.OPS_YYYYWW BETWEEN '" & ASCMAIN1.Week_Calc(RYW, -26 + 1) & "' AND '" & RYW & "' AND RSTRETL1.QTY_SOLD > 0 THEN RSTRETL1.OPS_YYYYWW ELSE NULL END) WK26_START" & vbCrLf _
            & ", 0 WKS" & vbCrLf _
            & sqlD _
            & " from RSTRETL1,ARTCUST2" & vbCrLf _
            & " where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO (+) = RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & sqlRange & vbCrLf _
            & IIf(initialize, " and ROWNUM < 1" & vbCrLf, "") _
            & " group by RSTRETL1.ITEM_CODE" & vbCrLf

        Dim sqlZ As String = Replace(Replace(sqlX, _
                                     "Select RSTRETL1.ITEM_CODE", _
                                     "Select RSTRETL1.ITEM_CODE,NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?') CUST_STORE_CLASS_CODE"), _
                                 "group by RSTRETL1.ITEM_CODE", _
                                 "group by RSTRETL1.ITEM_CODE,NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?')")

        Dim sqlC As String = Replace(Replace(sqlX, _
                                     "Select RSTRETL1.ITEM_CODE", _
                                     "Select NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?') CUST_STORE_CLASS_CODE"), _
                                 "group by RSTRETL1.ITEM_CODE", _
                                 "group by NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?')")

        ASCMAIN1.sql = "SELECT X.*" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
            & " from ICTITEM1,ICTSIZEN,ICTCOLL1,(" & vbCrLf _
            & sqlX _
            & ") X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
        sqlRSTCSKUW = ASCMAIN1.sql

        sqlZ = "SELECT X.*" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
            & " from ICTITEM1,ICTSIZEN,ICTCOLL1,(" & vbCrLf _
            & sqlZ _
            & ") X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"


        If initialize Then
            RSTCSKUW = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTCSKUW & " Add Primary Key (ITEM_CODE)")

            ASCMAIN1.sql = sqlZ
            RSTCSKUZ = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTCSKUZ & " Add Primary Key (ITEM_CODE,CUST_STORE_CLASS_CODE)")

            ASCMAIN1.sql = sqlC
            RSTCSKUC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTCSKUC & " Add Primary Key (CUST_STORE_CLASS_CODE)")

        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTCSKUW)
            ASCDATA1.ExecuteSQL("Insert into " & RSTCSKUW & " " & ASCMAIN1.sql)

            If MENU_ITEM_OBJECT = "RSFCSKUW" Then
                If chkUSEPLAN.Checked Then
                    Dim sqlPU As String = ""
                    Dim sqlPU1 As String = ""
                    Dim sqlPU2 As String = ""

                    Dim sql0 As String = ""
                    iPIVOT_CODE = -1
                    For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                        iPIVOT_CODE += 1
                        Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                        sql0 &= ", SLS_QTY_LY" & sfx & " = 0"

                        sqlPU &= "" _
                            & ", SUM (CASE WHEN RSTBUDRU.OPS_YYYYWW = '" & RYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTBUDRU.BUDGET ELSE 0 END) SLS_QTY_LY" & sfx & vbCrLf

                        sqlPU1 &= ", SLS_QTY_LY" & sfx
                        sqlPU2 &= ", SLS_QTY_LY" & sfx & " = R1.SLS_QTY_LY" & sfx

                    Next
                    ASCMAIN1.sql = "Update " & RSTCSKUW & " Set " & Mid(sql0, 2)
                    ASCDATA1.ExecuteSQL()
                    ASCMAIN1.sql = "Update " & RSTCSKUW & " Set " & Replace(Mid(sql0, 2), "SLS_QTY_LY", "SLS_AMT_LY")
                    ASCDATA1.ExecuteSQL()


                    Dim sqlP As String = "" _
                        & "Select RSTBUDRU.ITEM_CODE" _
                        & sqlPU _
                        & " from RSTBUDRU,ARTCUST2" & vbCrLf _
                        & " where RSTBUDRU.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & "   and ARTCUST2.CUST_CODE (+) = RSTBUDRU.CUST_CODE" & vbCrLf _
                        & "   and ARTCUST2.CUST_STORE_NO (+) = RSTBUDRU.CUST_STORE_NO" & vbCrLf _
                        & Replace(sqlRange, "RSTRETL1", "RSTBUDRU") & vbCrLf _
                        & " group by RSTBUDRU.ITEM_CODE" & vbCrLf


                    Dim sqlPP = "SELECT X.*" & vbCrLf _
                        & ", ICTITEM1.ITEM_DESC, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                        & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
                        & " from ICTITEM1,ICTSIZEN,ICTCOLL1,(" & vbCrLf _
                        & sqlP _
                        & ") X" & vbCrLf _
                        & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
                        & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
                        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"

                    Dim sqlI As String = ",ITEM_CODE,ITEM_DESC,HC_CODE,BRAND_CODE,ITEM_RETAIL_PRICE,ITEM_UPC_CODE,NRF_SIZE_DESC"
                    ASCMAIN1.sql = "" _
                        & "Begin" & vbCrLf _
                        & " Declare Cursor C1 is " & sqlPP & ";" & vbCrLf _
                        & " Begin" & vbCrLf _
                        & "  For R1 in C1 Loop" & vbCrLf _
                        & "   Update " & RSTCSKUW & " Set " & Mid(sqlPU2, 2) & vbCrLf _
                        & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                        & "   If SQL%NOTFOUND then" & vbCrLf _
                        & "    Insert into " & RSTCSKUW & vbCrLf _
                        & " (" & Mid(sqlI, 2) & sqlPU1 & ") Values (" & Mid(Replace(sqlI, ",", ",R1."), 2) & Replace(sqlPU1, ", ", ",R1.") & ");" & vbCrLf _
                        & "   End If;" & vbCrLf _
                        & "  End Loop;" & vbCrLf _
                        & " End;" & vbCrLf _
                        & "End;"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = Replace(Replace(ASCMAIN1.sql, "SLS_QTY_LY", "SLS_AMT_LY"), "RSTBUDRU", "RSTBUDRS")
                    ASCDATA1.ExecuteSQL()

                End If

            End If
          
            ASCDATA1.ExecuteSQL("Delete from " & RSTCSKUW & " where NVL(SLS_QTY_TY,0) = 0 and NVL(SLS_QTY_LY,0) = 0 and NVL(SLS_AMT_TY,0) = 0 and NVL(SLS_AMT_LY,0) = 0 and NVL(OHDC,0) = 0 and NVL(OHST,0) = 0 and NVL(SLS_L26,0) = 0")

            ASCMAIN1.sql = "Update " & RSTCSKUW & " RSTCSKUW SET WKS = (SELECT WKS FROM " & vbCrLf _
                & "(Select ROWNUM WKS, YYYYWW from GLTPARM3 where YYYYWW between '" & ASCMAIN1.Week_Calc(RYW, -26 + 1) & "' AND '" & RYW & "') " & vbCrLf _
                & " where YYYYWW = RSTCSKUW.WK26_START) " & vbCrLf _
                & " where WK26_START IS NOT NULL"
            ASCDATA1.ExecuteSQL()

            If MENU_ITEM_OBJECT = "RSFCSKUM" Then
                ASCDATA1.ExecuteSQL("Truncate Table " & RSTCSKUZ)
                ASCDATA1.ExecuteSQL("Insert into " & RSTCSKUZ & " " & sqlZ)
                ASCDATA1.ExecuteSQL("Delete from " & RSTCSKUZ & " where NVL(SLS_QTY_TY,0) = 0 and NVL(SLS_QTY_LY,0) = 0 and NVL(SLS_AMT_TY,0) = 0 and NVL(SLS_AMT_LY,0) = 0 and NVL(OHDC,0) = 0 and NVL(OHST,0) = 0 and NVL(SLS_L26,0) = 0")
                ASCDATA1.ExecuteSQL("Delete from " & RSTCSKUZ & " where ITEM_CODE in (Select ITEM_CODE from " & RSTCSKUZ & " minus Select ITEM_CODE from " & RSTCSKUW & ")")

                ASCDATA1.ExecuteSQL("Truncate Table " & RSTCSKUC)
                ASCDATA1.ExecuteSQL("Insert into " & RSTCSKUC & " " & Replace(sqlC, " where ", " where ITEM_CODE in (Select Distinct ITEM_CODE from " & RSTCSKUZ & ") and "))
                ASCDATA1.ExecuteSQL("Delete from " & RSTCSKUC & " where NVL(SLS_QTY_TY,0) = 0 and NVL(SLS_QTY_LY,0) = 0 and NVL(SLS_AMT_TY,0) = 0 and NVL(SLS_AMT_LY,0) = 0 and NVL(OHDC,0) = 0 and NVL(OHST,0) = 0 and NVL(SLS_L26,0) = 0")
                'ASCDATA1.ExecuteSQL("Delete from " & RSTCSKUC & " where ITEM_CODE in (Select ITEM_CODE from " & RSTCSKUC & " minus Select ITEM_CODE from " & RSTCSKUW & ")")







                If chkUSEPLAN.Checked Then
                    For Each TABLE_NAME As String In New String() {RSTCSKUZ, RSTCSKUC}

                        Dim sqlPU As String = ""
                        Dim sqlPU1 As String = ""
                        Dim sqlPU2 As String = ""

                        Dim sql0 As String = ""
                        iPIVOT_CODE = -1
                        For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                            iPIVOT_CODE += 1
                            Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                            sql0 &= ", SLS_QTY_LY" & sfx & " = 0"

                            'sqlPU &= "" _
                            '    & ", SUM (CASE WHEN RSTBUDRU.OPS_YYYYWW = '" & RYW & "' " & IIf(iPIVOT_CODE = 0, "", " AND ARTCUST2.CUST_STORE_CLASS_CODE = '" & PIVOT_CODE & "'") & " THEN RSTBUDRU.BUDGET ELSE 0 END) SLS_QTY_LY" & sfx & vbCrLf

                            Dim X As String = "CASE WHEN GLTPARM3.YYYYPP BETWEEN '000001' AND '000002' THEN RSTBUDRU.BUDGET ELSE 0 END"
                            If iPIVOT_CODE = 0 Then
                                X = Replace(X, "000001", "000000")
                                TY_YP = RYP
                            Else
                                TY_YP = ASCMAIN1.Period_Calc(RYP, -12 + iPIVOT_CODE)
                            End If

                            sqlPU &= "" _
                                & ", SUM (" & Replace(Replace(Replace(Replace(X, "000000", RYP_01), "000001", TY_YP), "000002", TY_YP), "XXX", "QTY_SOLD") & ") SLS_QTY_LY" & sfx & vbCrLf
                            sqlPU1 &= ", SLS_QTY_LY" & sfx
                            sqlPU2 &= ", SLS_QTY_LY" & sfx & " = R1.SLS_QTY_LY" & sfx

                        Next
                        ASCMAIN1.sql = "Update " & TABLE_NAME & " Set " & Mid(sql0, 2)
                        ASCDATA1.ExecuteSQL()
                        ASCMAIN1.sql = "Update " & TABLE_NAME & " Set " & Replace(Mid(sql0, 2), "SLS_QTY_LY", "SLS_AMT_LY")
                        ASCDATA1.ExecuteSQL()


                        Dim sqlP As String = "" _
                            & "Select RSTBUDRU.ITEM_CODE" _
                            & sqlPU _
                            & " from RSTBUDRU,ARTCUST2,GLTPARM3" & vbCrLf _
                            & " where RSTBUDRU.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and ARTCUST2.CUST_CODE (+) = RSTBUDRU.CUST_CODE" & vbCrLf _
                            & "   and ARTCUST2.CUST_STORE_NO (+) = RSTBUDRU.CUST_STORE_NO" & vbCrLf _
                            & "   and GLTPARM3.YYYYWW = RSTBUDRU.OPS_YYYYWW" & vbCrLf _
                            & Replace(Replace(sqlRange, "RSTRETL1", "RSTBUDRU"), "RSTBUDRU.OPS_YYYYPP", "GLTPARM3.YYYYPP") & vbCrLf _
                            & " group by RSTBUDRU.ITEM_CODE" & vbCrLf

                        Dim sqlPP = "SELECT X.*" & vbCrLf _
                            & ", ICTITEM1.ITEM_DESC, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
                            & " from ICTITEM1,ICTSIZEN,ICTCOLL1,(" & vbCrLf _
                            & sqlP _
                            & ") X" & vbCrLf _
                            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
                            & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
                            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"


                        Dim sqlI As String = ",ITEM_CODE,ITEM_DESC,HC_CODE,BRAND_CODE,ITEM_RETAIL_PRICE,ITEM_UPC_CODE,NRF_SIZE_DESC"

                        If TABLE_NAME = RSTCSKUC Then
                            sqlPP = Replace(Replace(sqlP, _
                               "Select RSTBUDRU.ITEM_CODE", _
                               "Select NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?') CUST_STORE_CLASS_CODE"), _
                           "group by RSTBUDRU.ITEM_CODE", _
                           "group by NVL(ARTCUST2.CUST_STORE_CLASS_CODE,'?')")

                            sqlI = ",CUST_STORE_CLASS_CODE"
                        End If


                        ASCMAIN1.sql = "" _
                            & "Begin" & vbCrLf _
                            & " Declare Cursor C1 is " & sqlPP & ";" & vbCrLf _
                            & " Begin" & vbCrLf _
                            & "  For R1 in C1 Loop" & vbCrLf _
                            & "   Update " & TABLE_NAME & " Set " & Mid(sqlPU2, 2) & vbCrLf _
                            & IIf(TABLE_NAME = RSTCSKUC, _
                                  "    where CUST_STORE_CLASS_CODE = R1.CUST_STORE_CLASS_CODE;", _
                                  "    where ITEM_CODE = R1.ITEM_CODE;") & vbCrLf _
                            & "   If SQL%NOTFOUND then" & vbCrLf _
                            & "    Insert into " & TABLE_NAME & vbCrLf _
                            & " (" & Mid(sqlI, 2) & sqlPU1 & ") Values (" & Mid(Replace(sqlI, ",", ",R1."), 2) & Replace(sqlPU1, ", ", ",R1.") & ");" & vbCrLf _
                            & "   End If;" & vbCrLf _
                            & "  End Loop;" & vbCrLf _
                            & " End;" & vbCrLf _
                            & "End;"
                        ASCDATA1.ExecuteSQL()

                        ASCMAIN1.sql = Replace(Replace(ASCMAIN1.sql, "SLS_QTY_LY", "SLS_AMT_LY"), "RSTBUDRU", "RSTBUDRS")
                        ASCDATA1.ExecuteSQL()
                    Next

                End If

            End If

        End If

    End Sub

    Private Sub btnAddItems_Click(sender As Object, e As EventArgs) Handles btnAddItems.Click

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ITEM_CODE")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True

            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Dim ITEM_CODEs As String = Mid(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), "','") & "'", 3)
                Dim sql As String = "Select Distinct ITEM_CODE from ICTITEM1 where ITEM_CODE in (" & ITEM_CODEs & ")"
                Add_Items(sql)

                For Each ITEM_CODE As String In Split(Mid(ASCMAIN1.CodeSelector.SelectedCodes0, 2), Chr(0))
                    Dim rowRSTCSKU1 As DataRow = dst.Tables("RSTCSKU1").Rows.Find(New String() {CUST_CODE, ITEM_CODE})
                    rowRSTCSKU1.Item("SEL") = "1"
                Next
            End If
        End If
    End Sub

    Private Sub btnAddSThru_Click(sender As Object, e As EventArgs) Handles btnAddSThru.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Making List of Items Sold-Thru in week " & RYW)

        Dim sql As String = "Select Distinct ITEM_CODE from RSTRETL1 where CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYWW = '" & RYW & "'"
        Add_Items(sql)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Add_Items(sql As String)
        ASCMAIN1.sql = "Select '" & CUST_CODE & "' CUST_CODE, X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_UPC_CODE" & vbCrLf _
            & ", ICTCOLL1.HC_CODE, ICTSIZEN.NRF_SIZE_DESC" & vbCrLf _
            & " from (" & Sql & ") X,ICTITEM1,ICTSIZEN,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTSIZEN.NRF_SIZE_CODE = ICTITEM1.NRF_SIZE_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            If dst.Tables("RSTCSKU1").Rows.Find(New String() {CUST_CODE, ITEM_CODE}) Is Nothing Then
                Dim rowRSTCSKU1 As DataRow = dst.Tables("RSTCSKU1").Rows.Add(row.ItemArray)
                Get_1st_Last_Sale_YW_OH(rowRSTCSKU1)
            End If
        Next
    End Sub

    Sub Get_1st_Last_Sale_YW_OH(row As DataRow)

        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        Dim rowRSTRETLX As DataRow = dst.Tables("RSTRETLX").Rows.Find(ITEM_CODE)

        If rowRSTRETLX IsNot Nothing Then
            row.Item("YW_MIN") = rowRSTRETLX.Item("YW_MIN")
            row.Item("YW_MAX") = rowRSTRETLX.Item("YW_MAX")
            row.Item("QTY_ONH") = rowRSTRETLX.Item("QTY_ONH")
        End If

    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdRSTCSKUW"
                Dim KEY As String = summarySettings.Key
                If KEY.Contains("VARPCT") Then
                    Dim TY As String = Replace(KEY, "VARPCT", "TY")
                    Dim LY As String = Replace(KEY, "VARPCT", "LY")
                    TOTALS.Add(TY, 0)
                    TOTALS.Add(LY, 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS(LY) <> 0 Then CustomValue = TOTALS(TY) / TOTALS(LY)

                ElseIf KEY = "" Then
                    Stop
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Public Overrides Function CustomStringSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As String, _
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdRSTCSKUW"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
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
                If KEY.Contains("VARPCT") Then
                    Dim TY As String = Replace(KEY, "VARPCT", "TY")
                    Dim LY As String = Replace(KEY, "VARPCT", "LY")
                    TOTALS(TY) += CStr(grow2.Cells(TY).Value & "")
                    TOTALS(LY) += CStr(grow2.Cells(LY).Value & "")

                ElseIf KEY = "" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Sub Set_COMP()
        If MENU_ITEM_OBJECT = "RSFCSKUM" Then

            'Dim dvw As DataView = dst.Tables("RSTCSKUZ").DefaultView
            'Dim SQL As String = ""
            'If optCOMP.Value = "C" Then
            '    SQL = "COMP = '1'"
            'ElseIf optCOMP.Value = "N" Then
            '    SQL = "COMP = '0'"
            'End If
            'dvw.RowFilter = SQL
            grdRSTCSKUW.DisplayLayout.Override.RowFilterMode = UltraWinGrid.RowFilterMode.AllRowsInBand
            With grdRSTCSKUW.DisplayLayout.Bands(1)
                .ColumnFilters.ClearAllFilters()
                If optCOMP.Value = "C" Then
                    .ColumnFilters("COMP").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, 1)
                    '.ColumnFilters("NRF_SIZE_DESC").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.StartsWith, "200 ML")
                ElseIf optCOMP.Value = "N" Then
                    '     .ColumnFilters("OPS_YYYYWW").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, OPS_YYYYWW)

                    '.ColumnFilters("NRF_SIZE_DESC").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.StartsWith, "20 ML")
                    .ColumnFilters("COMP").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, 0)
                End If
            End With

            grdRSTCSKUW.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)
        Else
            Dim dvw As DataView = DirectCast(grdRSTCSKUW.DataSource, DataTable).DefaultView
            Dim SQL As String = ""
            If optCOMP.Value = "C" Then
                SQL = "COMP = '1'"
            ElseIf optCOMP.Value = "N" Then
                SQL = "COMP = '0'"
            End If
            dvw.RowFilter = SQL

        End If
    End Sub

    Private Sub optCOMP_ValueChanged(sender As Object, e As EventArgs) Handles optCOMP.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_COMP()
    End Sub

    Sub Rank_SKUs()

        Dim iPIVOT_CODE As Integer = 0

        For Each row As DataRow In dst.Tables("RSTCSKUW").Select("")
            iPIVOT_CODE = -1
            For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                iPIVOT_CODE += 1
                Dim C As String
                C = "SLS_AMT_RANK" & IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                row.Item(C) = DBNull.Value
                C = "SLS_QTY_RANK" & IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                row.Item(C) = DBNull.Value
            Next
        Next

        iPIVOT_CODE = -1
        For Each PIVOT_CODE In PIVOT_CODEs
            iPIVOT_CODE += 1
            Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
            For Each COL As String In New String() {"SLS_AMT", "SLS_QTY"}
                Dim C As String = COL & "_RANK" & sfx
                Dim RANK_BY As String = COL & "_TY" & sfx
                Dim RANK = 0
                For Each row As DataRow In dst.Tables("RSTCSKUW").Select(RANK_BY & ">0", RANK_BY & " DESC")
                    RANK += 1
                    row.Item(C) = RANK
                Next
            Next
        Next


        For Each rowRSTCSKUC As DataRow In dst.Tables("RSTCSKUC").Select("")
            Dim CUST_STORE_CLASS_CODE As String = rowRSTCSKUC.Item("CUST_STORE_CLASS_CODE") & ""
            Dim sqlx As String = "CUST_STORE_CLASS_CODE = '" & CUST_STORE_CLASS_CODE & "'"

            For Each row As DataRow In dst.Tables("RSTCSKUZ").Select(sqlx)
                iPIVOT_CODE = -1
                For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                    iPIVOT_CODE += 1
                    Dim C As String
                    C = "SLS_AMT_RANK" & IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                    row.Item(C) = DBNull.Value
                    C = "SLS_QTY_RANK" & IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                    row.Item(C) = DBNull.Value
                Next
            Next

            iPIVOT_CODE = -1
            For Each PIVOT_CODE In PIVOT_CODEs
                iPIVOT_CODE += 1
                Dim sfx As String = IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                For Each COL As String In New String() {"SLS_AMT", "SLS_QTY"}
                    Dim C As String = COL & "_RANK" & sfx
                    Dim RANK_BY As String = COL & "_TY" & sfx
                    Dim RANK = 0
                    For Each row As DataRow In dst.Tables("RSTCSKUZ").Select(sqlx & " and " & RANK_BY & ">0", RANK_BY & " DESC")
                        RANK += 1
                        row.Item(C) = RANK
                    Next
                Next
            Next
        Next

    End Sub

    Sub Get_PIVOT_CODEs()

        PIVOT_CODEs.Clear()
        PIVOT_CODEs.Add("", "Totals")

        Select Case MENU_ITEM_OBJECT
            Case "RSFCSKUW"
                ASCMAIN1.sql = "Select CUST_STORE_CLASS_CODE, CUST_STORE_CLASS_DESC from ARTCLAS2"
                For Each rowARTCLAS2 As DataRow In ASCDATA1.GetDataTable.Select("", "CUST_STORE_CLASS_CODE")
                    Dim CUST_STORE_CLASS_CODE As String = rowARTCLAS2.Item("CUST_STORE_CLASS_CODE")
                    Dim CUST_STORE_CLASS_DESC As String = rowARTCLAS2.Item("CUST_STORE_CLASS_DESC")
                    PIVOT_CODEs.Add(CUST_STORE_CLASS_CODE, CUST_STORE_CLASS_DESC)
                Next
            Case "RSFCSKUM"
                ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" & vbCrLf _
                    & " where OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(RYP, -11) & "' and '" & RYP & "'"
                For Each rowGLTPARM2 As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP")
                    Dim OPS_YYYYPP As String = rowGLTPARM2.Item("OPS_YYYYPP")
                    Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                    PIVOT_CODEs.Add(OPS_YYYYPP, LEGEND)
                Next
        End Select
    End Sub

    Sub Format_grdRSTCSKU1()

        With grdRSTCSKU1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key <> "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If

                GCOL.Header.Appearance.BackColor = Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        Create_Summary(grdRSTCSKU1, "ITEM_CODE", "Count")
        Create_Summary(grdRSTCSKU1, "SEL")
    End Sub

    Sub Format_grdRSTCSKUX()
        With grdRSTCSKUX.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add()
            G.Header.Caption = "Customer / Store Information"
            G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_STORE_NO", "CUST_STORE_CLASS_CODE"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Hidden = False
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                    If COLUMN_NAME = "CUST_STORE_CLASS_CODE" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Chartreuse
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                End With
            Next

            With .Columns("CUST_CODE")
                .Header.Caption = "Customer"
                .Width = 90
            End With
            With .Columns("CUST_STORE_NO")
                .Header.Caption = "Store No"
                .Width = 70
            End With
            With .Columns("CUST_STORE_CLASS_CODE")
                .Header.Caption = "Class"
                .Width = 50
            End With

            .Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        End With

        Create_Summary(grdRSTCSKUX, "CUST_CODE", "Count")

        With grdRSTCSKUX.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
        End With
    End Sub

    Sub Format_grdRSTCSKUW()

        With grdRSTCSKUW.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False

            .AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
        End With

        If MENU_ITEM_OBJECT = "RSFCSKUM" Then
            ' grdRSTCSKUW.DisplayLayout.Bands(1).GroupHeadersVisible = False
            grdRSTCSKUW.DisplayLayout.Bands(1).ColHeadersVisible = True ' False
            grdRSTCSKUW.DisplayLayout.Bands(0).Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
            grdRSTCSKUW.DisplayLayout.Bands(1).Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
        End If

        For B As Integer = 0 To IIf(MENU_ITEM_OBJECT = "RSFCSKUM", 1, 0)
            With grdRSTCSKUW.DisplayLayout.Bands(B)
                For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                    '.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.Header.Appearance.BackColor = Drawing.Color.White
                    GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next

                Dim G As UltraWinGrid.UltraGridGroup

                If (MENU_ITEM_OBJECT = "RSFCSKUM" And B = 1) Or (MENU_ITEM_OBJECT <> "RSFCSKUM" And B = 0) Then
                    G = .Groups.Add("Item Master")
                    G.Header.Caption = "Item Master Data"
                    G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
                    G.Header.Appearance.BackColor = Drawing.Color.White
                    G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    For Each COLUMN_NAME As String In New String() _
                        {"ITEM_CODE", "ITEM_DESC", "HC_CODE", "BRAND_CODE", _
                         "ITEM_RETAIL_PRICE", "ITEM_UPC_CODE", "NRF_SIZE_DESC"} ', "COMP"
                        With .Columns(COLUMN_NAME)
                            .Group = G
                            .Hidden = False
                            .Header.Appearance.BackColor2 = Drawing.Color.Gold
                        End With
                    Next
                Else
                    G = .Groups.Add("Store Class")
                    'G.Header.Caption = "Item Master Data"
                    'G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
                    'G.Header.Appearance.BackColor = Drawing.Color.White
                    'G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_STORE_CLASS_CODE", _
                         "FILLER1", "FILLER7", "FILLER2", "FILLER3", "FILLER4", "FILLER5", "FILLER6"} ', "FILLER7"
                        If COLUMN_NAME.StartsWith("FILLER") Then
                            .Columns.Add(COLUMN_NAME)
                        End If
                        With .Columns(COLUMN_NAME)
                            .Group = G
                            .Hidden = False
                            .Header.Appearance.BackColor2 = Drawing.Color.Gold
                        End With
                    Next
                End If


                Dim iPIVOT_CODE As Integer = -1
                For Each PIVOT_CODE As String In PIVOT_CODEs.Keys
                    iPIVOT_CODE += 1
                    G = .Groups.Add(Format(iPIVOT_CODE, "000"))
                    G.Header.Caption = PIVOT_CODEs(PIVOT_CODE)
                    G.Header.Appearance.TextHAlign = HAlign.Center
                    G.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    G.Header.Appearance.BackColor = Drawing.Color.White
                    G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    For Each COLUMN_NAME As String In New String() {"SLS_QTY", "SLS_AMT"}
                        For Each Y As String In New String() {"TY", "LY", "VAR", "VARPCT", "RANK"}
                            Dim C As String = COLUMN_NAME & "_" & Y & IIf(iPIVOT_CODE = 0, "", "_" & Format(iPIVOT_CODE, "000"))
                            Dim PFX As String = IIf(COLUMN_NAME = "SLS_QTY", "#", "$")
                            With .Columns(C)
                                .Group = G
                                .Hidden = False
                                If COLUMN_NAME = "SLS_QTY" And Y = "RANK" Then
                                    .Hidden = True
                                End If

                                If COLUMN_NAME = "SLS_QTY" Then
                                    .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                                Else
                                    .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                                End If

                                If Y = "VARPCT" Then
                                    .Header.Caption = "+/-%"
                                    .Format = "#,##0.0%"
                                    .Width = 60
                                Else
                                    If Y = "VAR" Then
                                        If COLUMN_NAME = "SLS_QTY" Then
                                            .Header.Caption = "+/-#"
                                        Else
                                            .Header.Caption = "+/-$"
                                        End If
                                    Else
                                        If Y = "RANK" Then
                                            .Header.Caption = PFX & "Rank"
                                            .CellAppearance.BackColor = Color.LightGreen
                                        Else
                                            .Header.Caption = PFX & Y
                                        End If
                                    End If
                                    If Y = "RANK" Then
                                        .Format = "#,##0"
                                    Else
                                        .Format = PFX & "#,##0"

                                    End If
                                    .Width = 70
                                End If

                            End With
                            If Y = "RANK" Then
                            ElseIf Y = "VARPCT" Then
                                'If B = 0 Then Create_Summary(grdRSTCSKUW, C, "Custom", .Key)
                                Create_Summary(grdRSTCSKUW, C, "Custom", .Key)
                            Else
                                'If B = 0 Then Create_Summary(grdRSTCSKUW, C, , .Key)
                                Create_Summary(grdRSTCSKUW, C, , .Key)
                            End If
                        Next
                    Next
                Next

                G = .Groups.Add("Inventory")
                G.Header.Caption = "Inventory @" & cbeYW.Text
                G.Header.Appearance.TextHAlign = HAlign.Center
                G.Header.Appearance.BackColor2 = Drawing.Color.Chartreuse
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                For Each COLUMN_NAME As String In New String() {"OH", "OHST", "OHDC"}
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .Hidden = False
                        .Format = "#,##0"
                        .Width = 80
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        If B = 0 Then Create_Summary(grdRSTCSKUW, COLUMN_NAME, , .Band.Key)
                    End With
                Next


                G = .Groups.Add("WOS")
                G.Header.Caption = "Weeks of Supply"
                G.Header.Appearance.TextHAlign = HAlign.Center
                G.Header.Appearance.BackColor2 = Drawing.Color.Violet
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                For Each COLUMN_NAME As String In New String() {"SLS_L26", "WKS", "WOS", "WOS_ST", "WOS_DC"}
                    With .Columns(COLUMN_NAME)
                        .Group = G
                        .Hidden = False
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        If COLUMN_NAME = "SLS_L26" Or COLUMN_NAME = "WKS" Then
                            .Format = "#,##0"
                        Else
                            .Format = "#,##0.0"
                        End If
                        .Width = 60
                        ' Create_Summary(grdRSTCSKUW, COLUMN_NAME)
                    End With
                Next

                If (MENU_ITEM_OBJECT = "RSFCSKUM" And B = 1) Or (MENU_ITEM_OBJECT <> "RSFCSKUM" And B = 0) Then
                    With .Columns("ITEM_CODE")
                        .Header.Caption = "Item"
                        .Width = 100
                    End With
                    With .Columns("ITEM_DESC")
                        .Header.Caption = "Description"
                        .Width = 120
                    End With
                    With .Columns("HC_CODE")
                        .Header.Caption = "HC"
                        .Width = 50
                    End With
                    With .Columns("BRAND_CODE")
                        .Header.Caption = "Brand"
                        .Width = 50
                    End With
                    With .Columns("ITEM_RETAIL_PRICE")
                        .Header.Caption = "Retail"
                        .Format = "###.00"
                        .Width = 50
                    End With
                    With .Columns("ITEM_UPC_CODE")
                        .Header.Caption = "UPC"
                        .Width = 40
                    End With
                    With .Columns("NRF_SIZE_DESC")
                        .Header.Caption = "Size"
                        .Width = 60
                    End With
                    'With .Columns("COMP")
                    '    .Header.Caption = "Comp"
                    '    .Width = 60
                    'End With
                Else
                    With .Columns("CUST_STORE_CLASS_CODE")
                        .Header.Caption = "Item"
                        .Width = 120
                    End With
                    With .Columns("FILLER1")
                        .Header.Caption = "Description"
                        .Width = 60
                    End With
                    With .Columns("FILLER7")
                        .Header.Caption = ""
                        .Width = 60
                    End With
                    With .Columns("FILLER2")
                        .Header.Caption = "HC"
                        .Width = 50
                    End With

                    With .Columns("FILLER3")
                        .Header.Caption = "Brand"
                        .Width = 50
                    End With

                    With .Columns("FILLER4")
                        .Header.Caption = "Retail"
                        .Width = 50
                    End With

                    With .Columns("FILLER5")
                        .Header.Caption = "UPC"
                        .Width = 40
                    End With

                    With .Columns("FILLER6")
                        .Header.Caption = "Size"
                        .Width = 60
                    End With

                    'With .Columns("FILLER7")
                    '    .Header.Caption = "Comp"
                    '    .Width = 60
                    'End With
                End If


                With .Columns("OH")
                    .Header.Caption = "Total"
                    .CellAppearance.BackColor = Color.LightGray
                End With
                With .Columns("OHST")
                    .Header.Caption = "OH Str"
                End With
                With .Columns("OHDC")
                    .Header.Caption = "OH DC"
                End With

                With .Columns("SLS_L26")
                    .Header.Caption = "Sls26w"
                    .CellAppearance.BackColor = Color.LightGray
                End With
                With .Columns("WOS")
                    .Header.Caption = "W/S"
                End With
                With .Columns("WOS_ST")
                    .Header.Caption = "W/S ST"
                End With
                With .Columns("WOS_DC")
                    .Header.Caption = "W/S DC"
                End With

                .Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
            End With
        Next

        If (MENU_ITEM_OBJECT = "RSFCSKUM") Then
            Create_Summary(grdRSTCSKUW, "ITEM_CODE", "Count", "RSTCSKUC_RSTCSKUZ")

        Else
            Create_Summary(grdRSTCSKUW, "ITEM_CODE", "Count")

        End If

        With grdRSTCSKUW.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
            If (MENU_ITEM_OBJECT = "RSFCSKUM") Then
                grdRSTCSKUW.DisplayLayout.Bands(1).Groups(0).Header.Fixed = True
            End If
        End With
    End Sub

    Private Sub btnExportByStore_Click(sender As Object, e As EventArgs) Handles btnExportByStore.Click

        ' Declare SSG Objects

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        'Dim rangeCopyFrom As SpreadsheetGear.IRange
        'Dim rangePaste_To As SpreadsheetGear.IRange

        ' Parameters

        Dim Start_Row As Integer = 5

        ' Save Workbook as FILENAME

        Dim FILENAME_TEMPLATE As String = ""
        Dim FILENAME_SOURCE As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & FILENAME_TEMPLATE
        Dim XLS_FILENAME As String = ""

        ASCMAIN1.Progress("Now Creating Custom XLS Workbook")
        If FILENAME_TEMPLATE = "" Then
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "Data"
            XLS_FILENAME = ASCMAIN1.Folders("Work") & XNO & ".xlsx"
            oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Else
            Dim success As Boolean = False
            Dim XLS_NO As Integer = 0

            Do Until success
                Try
                    XLS_NO += 1
                    XLS_FILENAME = ASCMAIN1.Folders("Work") & "Inventory_SellIn_SellThru_by_Store"
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                    FileCopy(FILENAME_SOURCE, XLS_FILENAME)
                    success = True

                Catch ex As Exception
                    ' Stop
                End Try
            Loop

            oWB = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)
            oSheet = oWB.Worksheets("Data")
        End If

        ' Worksheet Heading

        With oSheet.Cells(0, 0)
            ' .Value = Format(Now, "MM/dd/yyyy HH:mm")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .NumberFormat = "mm/dd/yy;@"
            .Value = Now
        End With
        With oSheet.Cells(0, 1)
            .Value = MENU_ITEM_OBJECT
        End With
        With oSheet.Cells(0, 2)
            .Value = ASCMAIN1.USER_ID
        End With
        With oSheet.Cells(1, 0)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Font.Name = "Times New Roman"
            ' .Name = "Verdana"
            .Value = MENU_ITEM_DESC
        End With
        With oSheet.Cells(2, 0)
            .Font.Color = SpreadsheetGear.Colors.Blue
            .Font.Size = 20
            .Value = "Inventory & Monthly Sell-In / Sell-Thru by Store"
        End With

        With oSheet.Cells(3, 0)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Value = "In Units"
        End With

        With oSheet.Cells(3, 1)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Value = cbeYW.Text
        End With

        With oSheet.Cells(1, 2)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Value = Absx1.txtFor("CUST_CODE").Text
        End With

        ' Prepare to Traverse Dataset

        Dim R As Integer = Start_Row
        oSheet.Cells(R - 1, 0, R, 0).EntireRow.NumberFormat = "@"

        Dim C As Integer = -1
        Dim COLs() As String = {"ITEM_CODE", "ITEM_DESC", "HC_CODE", "BRAND_CODE", "ITEM_RETAIL_PRICE", "ITEM_UPC_CODE", "NRF_SIZE_DESC"}
        For Each COLUMN_NAME As String In COLs
            C += 1
            With grdRSTCSKUW.DisplayLayout.Bands(1).Columns(COLUMN_NAME)
                oSheet.Cells(R, C).Value = .Header.Caption
                oSheet.Cells(R, C).EntireColumn.ColumnWidth = .Width / 10
                If dst.Tables("RSTCSKUW").Columns(COLUMN_NAME).DataType = GetType(System.String) Then
                    oSheet.Cells(R, C).EntireColumn.NumberFormat = "@"
                Else
                    If .Format <> "" Then
                        oSheet.Cells(R, C).EntireColumn.NumberFormat = .Format
                    End If
                End If
            End With
        Next

        oSheet.Cells(R, 0, R, COLs.Count - 1).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
        Prepare_Custom_XLS_Border(oSheet, R, 0, R, COLs.Count - 1)

        Dim CS As Integer = C ' STORES START HERE
        Dim ITEM_count As Integer = dst.Tables("RSTCSKUW").Rows.Count

        Dim CUST_STORE_NOs As New Dictionary(Of String, Integer)

        ' Total Cells

        oSheet.Cells(R + 1, C + 1, R + ITEM_count, C + 3 + 1).NumberFormat = "#,##0"
        oSheet.Cells(R + 1, C + 1, R + ITEM_count, C + 3 + 1).EntireColumn.ColumnWidth = 6

        oSheet.Cells(R - 1, C + 1).Value = "Total"
        oSheet.Cells(R - 1, C + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        oSheet.Cells(R - 1, C + 1).Interior.Color = SpreadsheetGear.Colors.PaleGreen
        oSheet.Cells(R - 1, C + 1, R - 1, C + 3 + 1).Merge()

        Prepare_Custom_XLS_Border(oSheet, R - 1, C + 1, R - 1, C + 3 + 1)
        Prepare_Custom_XLS_Border(oSheet, R + 1, C + 1, R + ITEM_count, C + 3 + 1)

        'oSheet.Cells(R - 1, C + 1, R - 1, C + 3).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        'oSheet.Cells(R - 1, C + 1, R - 1, C + 3).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke

        oSheet.Cells(R, C + 1, R, C + 3 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        C += 1 : oSheet.Cells(R, C).Value = "S-In"
        C += 1 : oSheet.Cells(R, C).Value = "Thru"
        C += 1 : oSheet.Cells(R, C).Value = "EOW"
        C += 1 : oSheet.Cells(R, C).Value = "Rank"


        For Each row As DataRow In ASCDATA1.SelectDistinct("RSTCSKUS", New String() {"CUST_STORE_NO"}).Select("")
            Dim CUST_STORE_NO As String = row.Item(0)
            CUST_STORE_NOs.Add(CUST_STORE_NO, C)
            oSheet.Cells(R + 1, C + 1, R + ITEM_count, C + 3).NumberFormat = "#,##0"
            oSheet.Cells(R + 1, C + 1, R + ITEM_count, C + 3).EntireColumn.ColumnWidth = 5

            oSheet.Cells(R - 1, C + 1).Value = CUST_STORE_NO
            oSheet.Cells(R - 1, C + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            oSheet.Cells(R - 1, C + 1).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
            oSheet.Cells(R - 1, C + 1, R - 1, C + 3).Merge()

            Prepare_Custom_XLS_Border(oSheet, R - 1, C + 1, R - 1, C + 3)
            Prepare_Custom_XLS_Border(oSheet, R + 1, C + 1, R + ITEM_count, C + 3)

            'oSheet.Cells(R - 1, C + 1, R - 1, C + 3).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            'oSheet.Cells(R - 1, C + 1, R - 1, C + 3).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke

            oSheet.Cells(R, C + 1, R, C + 3).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            C += 1 : oSheet.Cells(R, C).Value = "S-In"
            C += 1 : oSheet.Cells(R, C).Value = "Thru"
            C += 1 : oSheet.Cells(R, C).Value = "EOW"

        Next



        Dim SQL As String = ""
        If optCOMP.Value = "C" Then
            SQL = "COMP = '1'"
        ElseIf optCOMP.Value = "N" Then
            SQL = "COMP = '0'"
        End If

        For Each row As DataRow In dst.Tables("RSTCSKUW").Select(SQL, "ITEM_CODE")
            R += 1
            C = -1
            For Each COLUMN_NAME As String In COLs
                C += 1
                oSheet.Cells(R, C).Value = row.Item(COLUMN_NAME)
            Next
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim SQLW As String = "ITEM_CODE = '" & ITEM_CODE & "'"

            C += 1 : oSheet.Cells(R, C).Value = Val(dst.Tables("RSTCSKUS").Compute("SUM(SELLIN)", SQLW) & "")
            C += 1 : oSheet.Cells(R, C).Value = Val(dst.Tables("RSTCSKUS").Compute("SUM(SELLTHRU)", SQLW) & "")
            C += 1 : oSheet.Cells(R, C).Value = Val(dst.Tables("RSTCSKUS").Compute("SUM(ONHAND)", SQLW) & "")
            C += 1 : oSheet.Cells(R, C).Value = row.Item("SLS_QTY_RANK")


            For Each rowRSTCSKUS As DataRow In dst.Tables("RSTCSKUS").Select(SQLW, "CUST_STORE_NO")
                Dim CUST_STORE_NO As String = rowRSTCSKUS.Item("CUST_STORE_NO")
                C = CUST_STORE_NOs(CUST_STORE_NO)
                C += 1 : oSheet.Cells(R, C).Value = rowRSTCSKUS.Item("SELLIN")
                C += 1 : oSheet.Cells(R, C).Value = rowRSTCSKUS.Item("SELLTHRU")
                C += 1 : oSheet.Cells(R, C).Value = rowRSTCSKUS.Item("ONHAND")
            Next
        Next

        'oSheet.WindowInfo.DisplayGridlines = False


        ' Save Document and Show

        oWB.Save()
        Show_Document(XLS_FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Prepare_Custom_XLS_Border(oSheet As SpreadsheetGear.IWorksheet, R1 As Int64, C1 As Int64, R2 As Int64, C2 As Int64)
        With oSheet.Range(R1, C1, R2, C2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With
    End Sub
End Class