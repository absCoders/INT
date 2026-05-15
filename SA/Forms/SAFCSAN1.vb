Public Class SAFCSCAN1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim BRAND_CODE As String
    Dim rowICTBRAN1 As DataRow
    Dim sqlSATCSANX As String
    Dim CUST_STORE_CLASS_CODEs As New Dictionary(Of String, Integer)
    Dim i_to_CUST_STORE_CLASS_CODE As New Dictionary(Of Integer, String)
    Dim RYW As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SOFDRBRI" Then
            InquiryMode = True
        End If

        Create_WorkFile()

        With dst
            ASCMAIN1.sql = "Select * from SATCSAN1"
            Create_TDA(.Tables.Add, "SATCSAN1", "*")

            ASCMAIN1.sql = "Select * from ARTCUST2"
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, , , , "CUST_STORE_CUST_RANK_W,CUST_STORE_CUST_RANK_M")

            sqlSATCSANX = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.SELL_CODE, ARTCUST2.MALL_CODE, ARTCUST2.CUST_STORE_CUST_RANK_W, ARTCUST2.CUST_STORE_CUST_RANK_M" _
                & ", ARTCUST1.SREP_CODE" _
                & " from ARTCUST1, ARTCUST2" _
                & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE"
            ASCMAIN1.sql = sqlSATCSANX & " and ARTCUST2.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATCSANX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select * from ICTBRAN1"
            Create_TDA(.Tables.Add, "ICTBRAN1", "**", 0, False)
            .Tables("ICTBRAN1").Columns.Add("SEL")
            .Tables("ICTBRAN1").Columns("SEL").DefaultValue = "0"
            Fill_Records("ICTBRAN1")

            ASCMAIN1.sql = "Select * from ICTCOLL0"
            Create_TDA(.Tables.Add, "ICTCOLL0", "**", 0, False)
            Fill_Records("ICTCOLL0")

            Dim I As Integer = 0
            For Each rowICTCOLL0 As DataRow In dst.Tables("ICTCOLL0").Select("", "BRAND_CODE,HC_CODE")
                I += 1
                Dim DC As DataColumn = .Tables("SATCSANX").Columns.Add("C" & Format(I, "000"))
                DC.Caption = rowICTCOLL0.Item("HC_CODE")
                CUST_STORE_CLASS_CODEs.Add(rowICTCOLL0.Item("HC_CODE"), I)
                i_to_CUST_STORE_CLASS_CODE.Add(I, rowICTCOLL0.Item("HC_CODE"))
            Next
        End With

        grdICTBRAN1.DataSource = dst.Tables("ICTBRAN1")
        Sort_grdColumns(grdICTBRAN1, "BRAND_CODE")

        lblBRAND_CODE.Top = lblCUST_CODE.Top
        txtBRAND_CODE.Top = txtCUST_CODE.Top
        txtBRAND_NAME.Top = txtCUST_NAME.Top

        grdSATCSANX.DataSource = dst.Tables("SATCSANX")

        With grdSATCSANX.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add()
            G.Header.Caption = "Customer / Store Information"
            G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME In New String() {"CUST_CODE", "CUST_STORE_NO", "SREP_CODE", "SELL_CODE", "MALL_CODE", "CUST_STORE_CUST_RANK_W", "CUST_STORE_CUST_RANK_M"}
                With .Columns(COLUMN_NAME)
                    .Group = G
                    .Hidden = False
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                    If COLUMN_NAME = "CUST_STORE_CUST_RANK_W" Or COLUMN_NAME = "CUST_STORE_CUST_RANK_M" Then
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
            With .Columns("SREP_CODE")
                .Header.Caption = "S-In"
                .Width = 50
            End With
            With .Columns("SELL_CODE")
                .Header.Caption = "S-Thru"
                .Width = 50
            End With
            With .Columns("MALL_CODE")
                .Header.Caption = "Mall"
                .Width = 90
            End With
            With .Columns("CUST_STORE_CUST_RANK_W")
                .Header.Caption = "Rank" & vbCrLf & "W"
                .Width = 50
            End With
            With .Columns("CUST_STORE_CUST_RANK_M")
                .Header.Caption = "Rank" & vbCrLf & "M"
                .Width = 50
            End With
            Dim I As Integer = 0
            Dim B As Integer = 0
            For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("", "BRAND_CODE")
                B += 1
                Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
                G = .Groups.Add()
                G.Key = BRAND_CODE
                G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
                G.Header.Appearance.BackColor = Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                .Header.ToolTipText = rowICTBRAN1.Item("BRAND_CODE") & vbCrLf & rowICTBRAN1.Item("BRAND_NAME")

                If B Mod 2 = 0 Then
                    G.Header.Appearance.BackColor2 = Drawing.Color.Violet
                    G.Header.Appearance.BackColor = Drawing.Color.White
                Else
                    G.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    G.Header.Appearance.BackColor = Drawing.Color.White
                End If

                For Each rowICTCOLL0 As DataRow In dst.Tables("ICTCOLL0").Select("BRAND_CODE = '" & BRAND_CODE & "'", "HC_CODE")
                    I += 1
                    Dim COLUMN_NAME As String = "C" & Format(I, "000")
                    With .Columns(COLUMN_NAME)
                        .Header.Caption = rowICTCOLL0.Item("HC_CODE")
                        .Group = G
                        .Hidden = False
                        .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                        .Header.Appearance.TextHAlign = HAlign.Center
                        .Header.Appearance.TextVAlign = VAlign.Bottom
                        .CellAppearance.TextHAlign = HAlign.Center
                        .Header.Appearance.FontData.SizeInPoints = 8
                        .Header.TextOrientation = New TextOrientationInfo(90, TextFlowDirection.Horizontal)
                        .Width = 30
                        .Header.ToolTipText = rowICTCOLL0.Item("HC_CODE") & vbCrLf & rowICTCOLL0.Item("HC_NAME")
                        Create_Summary(grdSATCSANX, COLUMN_NAME, "Custom")
                    End With
                Next
            Next
            .Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        End With

        Create_Summary(grdSATCSANX, "CUST_CODE", "Count")

        With grdSATCSANX.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 3) & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(3)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"
                If optFilter.Value = "C" Then
                    If Absx1.txtFor("CUST_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Select a Customer Code"
                    Else
                        Validate_Code("CUST_CODE")
                        If EMsg = "" Then
                            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                            BRAND_CODE = ""
                            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                            If rowARTCUST1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                            End If
                        End If
                    End If
                Else
                    If Absx1.txtFor("BRAND_CODE").Text = "" Then
                        'EMsg &= vbCr & "You Must Select a Brand Code"
                        CUST_CODE = ""
                        BRAND_CODE = ""
                    Else
                        Validate_Code("BRAND_CODE")
                        If EMsg = "" Then
                            CUST_CODE = ""
                            BRAND_CODE = Absx1.txtFor("BRAND_CODE").Text
                            rowICTBRAN1 = LookUp("ICTBRAN1", BRAND_CODE)
                            If rowICTBRAN1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Brand Code " & BRAND_CODE
                            End If
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If optFilter.Value = "C" Then
                        If Not ASCMAIN1.Logical_Lock("SATAUTH1", "CUST_CODE:" & CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("SATAUTH1", "CUST_CODE:*") Then Exit Sub
                    Else
                        If Not ASCMAIN1.Logical_Lock("SATAUTH1", "CUST_CODE:*") Then Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update", "Save"

            Case "Load Ranks from XLS"


                For Each COLUMN_NAME As String In New String() {"SREP_CODE", "SELL_CODE", "MALL_CODE"}
                    grdSATCSANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                Next
                Excel_Import(grdSATCSANX)
                For Each COLUMN_NAME As String In New String() {"SREP_CODE", "SELL_CODE", "MALL_CODE"}
                    grdSATCSANX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
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

            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update", "Save"
                Update_Record()
                If eItemKey = "Update" Then
                    Mode_Settings(False)
                End If

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
                    .Items("Save").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not (EntryMode = "E")
                    .Items("Edit").Visible = (Not InquiryMode) And Not (EntryMode = "E")
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Save").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Load Ranks from XLS").Visible = ScreenMode And (Not InquiryMode And EntryMode <> "V") And (optFilter.Value = "C")

                End With
                '   .Groups("Period").Visible = ScreenMode
                '  .Groups("Period").Enabled = Not (EntryMode = "E")
                Set_Read_Only_for_ctl(cbeYP, (EntryMode = "E"))
                .Groups("Brands").Visible = ScreenMode And (optFilter.Value = "C" Or BRAND_CODE = "")
                .Groups("Copy Auths").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
            End With
        End If

        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                grdSATCSANX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSATCSANX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSATCSANX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdSATCSANX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSATCSANX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSATCSANX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl.Panel1Collapsed = ScreenMode
        splSATCSANX.Visible = ScreenMode

        splSATCSANX.Panel2Collapsed = True

        If ScreenMode Then
            Set_Groups()

            Dim W As Integer = 30
            If EntryMode = "V" Then W = 65
            For Each HC_CODE As String In CUST_STORE_CLASS_CODEs.Keys
                Dim C As Integer = CUST_STORE_CLASS_CODEs(HC_CODE)
                grdSATCSANX.DisplayLayout.Bands(0).Columns("C" & Format(C, "000")).Width = W
            Next
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SATCSANX", "ARTCUST2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If optFilter.Value = "B" Then
            ASCMAIN1.sql = sqlSATCSANX
            Fill_Records("SATCSANX", "", , ASCMAIN1.sql)
            grdSATCSANX.Text = "Store / Brand Matrix for " _
                & IIf(BRAND_CODE = "", "All Brands", BRAND_CODE) _
                & ", for All Customers"

            ASCMAIN1.sql = "Select * from SATAUTH1"
            If BRAND_CODE <> "" Then
                ASCMAIN1.sql &= " where HC_CODE in (Select HC_CODE from ICTCOLL0 where BRAND_CODE = '" & BRAND_CODE & "')"
            End If
            Fill_Records("SATAUTH1", "", , ASCMAIN1.sql)
        Else
            Fill_Records("SATCSANX", New String() {CUST_CODE})
            grdSATCSANX.Text = "Store / Brand Matrix for Customer " & CUST_CODE

            ASCMAIN1.sql = "Select * from SATAUTH1 where CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SATAUTH1", "", , ASCMAIN1.sql)

            Fill_Records("ARTCUST2", CUST_CODE)
        End If

        Dim BRAND_CODEs_In As New List(Of String)
        Dim HC_CODEs_In As New List(Of String)

        grdSATCSANX.SuspendSummaryUpdates()
        grdSATCSANX.SuspendLayout()

        For Each row As DataRow In dst.Tables("SATAUTH1").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim HC_CODE As String = row.Item("HC_CODE")
            Dim i As Integer = CUST_STORE_CLASS_CODEs(HC_CODE)
            Dim rowSATCSANX As DataRow = dst.Tables("SATCSANX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If row.Item("OPS_YYYYPP_OPENED") & "" <> "" Then
                If row.Item("OPS_YYYYPP_CLOSED") & "" <> "" Then
                    Dim YYYYMM As String = row.Item("OPS_YYYYPP_CLOSED")
                    Dim MMYY As String = Mid(YYYYMM, 5, 2) & "/" & Mid(YYYYMM, 3, 2)
                    rowSATCSANX.Item("C" & Format(i, "000")) = "C" & IIf(EntryMode = "V", ":" & MMYY, "")
                    If Not HC_CODEs_In.Contains(HC_CODE) Then HC_CODEs_In.Add(HC_CODE)
                Else
                    Dim YYYYMM As String = row.Item("OPS_YYYYPP_OPENED")
                    Dim MMYY As String = Mid(YYYYMM, 5, 2) & "/" & Mid(YYYYMM, 3, 2)
                    rowSATCSANX.Item("C" & Format(i, "000")) = "O" & IIf(EntryMode = "V", ":" & MMYY, "")
                    If Not HC_CODEs_In.Contains(HC_CODE) Then HC_CODEs_In.Add(HC_CODE)
                End If
            End If
        Next

        grdSATCSANX.ResumeSummaryUpdates(True)
        grdSATCSANX.ResumeLayout()

        For Each HC_CODE As String In HC_CODEs_In
            Dim rowICTCOLL0 As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
            Dim BRAND_CODE As String = rowICTCOLL0.Item("BRAND_CODE") & ""
            If BRAND_CODE <> "" AndAlso Not BRAND_CODEs_In.Contains(BRAND_CODE) Then
                BRAND_CODEs_In.Add(BRAND_CODE)
            End If
        Next

        Sort_grdColumns(grdSATCSANX, "CUST_CODE,CUST_STORE_NO")

        For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("")
            Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
            If Me.BRAND_CODE <> "" Then
                rowICTBRAN1.Item("SEL") = IIf(BRAND_CODE = Me.BRAND_CODE, "1", "0")
            Else
                rowICTBRAN1.Item("SEL") = IIf(rowICTBRAN1.Item("BRAND_STATUS") & "" = "I", "0", IIf(optFilter.Value = "B" Or BRAND_CODEs_In.Contains(BRAND_CODE), "1", "0"))
            End If
        Next

        dst.Tables("SATCSANX").AcceptChanges()


        EnforceConstraints(True)
        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-composing Authorization Records")

        Try
            BeginTrans()
            Update_Record_TDA("SATAUTH1")
            If optFilter.Value = "C" Then Update_Record_TDA("ARTCUST2")


            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSANX, "SS", "Show Filter", "Show GroupBox")
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
                Case "grdSATCSANX"

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
            Case "CUST_CODE", "BRAND_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE", "BRAND_CODE"
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

    Private Sub grdICTBRAN1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdICTBRAN1.AfterRowUpdate
        Set_Groups()
    End Sub

    Sub Set_Groups()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Screen")
        For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("")
            Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
            grdSATCSANX.DisplayLayout.Bands(0).Groups(BRAND_CODE).Hidden = (rowICTBRAN1.Item("SEL") & "" <> "1")
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSATCSANX_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSATCSANX.ClickCell
        If e.Cell.Row.IsDataRow Then
            Dim CUST_CODE As String = e.Cell.Row.Cells("CUST_CODE").Value
            Dim CUST_STORE_NO As String = e.Cell.Row.Cells("CUST_STORE_NO").Value
            Dim HC_CODE As String = e.Cell.Column.Header.Caption

            Dim rowSATAUTH1 As DataRow = dst.Tables("SATAUTH1").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, HC_CODE})
            If rowSATAUTH1 IsNot Nothing Then
                If rowSATAUTH1.Item("OPS_YYYYPP_OPENED") & "" <> "" Then
                    e.Cell.ToolTipText = "Opened " & rowSATAUTH1.Item("OPS_YYYYPP_OPENED")
                    If rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") & "" <> "" Then
                        e.Cell.ToolTipText &= "; Closed " & rowSATAUTH1.Item("OPS_YYYYPP_CLOSED")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub grdSATCSANX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATCSANX.InitializeRow

    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSATCSANX"
                Dim KEY As String = summarySettings.Key
                If KEY.StartsWith("C") And KEY.Length = 4 Then
                    TOTALS.Add("O", 0)
                    CustomSummary_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("O") <> 0 Then CustomValue = TOTALS("O")

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
            Case "grdSATCSANX"
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
                If KEY.StartsWith("C") And KEY.Length = 4 Then
                    If CStr(grow2.Cells(KEY).Value & "").StartsWith("O") Then
                        TOTALS("O") += 1
                    End If

                ElseIf KEY = "" Then
                    '  TOTALS(KEY) = "Totals"
                End If
            End If
        Next
    End Sub

    Private Sub optFilter_ValueChanged(sender As Object, e As EventArgs) Handles optFilter.ValueChanged
        lblCUST_CODE.Visible = (optFilter.Value = "C")
        txtCUST_CODE.Visible = (optFilter.Value = "C")
        txtCUST_NAME.Visible = (optFilter.Value = "C")

        lblBRAND_CODE.Visible = (optFilter.Value = "B")
        txtBRAND_CODE.Visible = (optFilter.Value = "B")
        txtBRAND_NAME.Visible = (optFilter.Value = "B")

        lblBRAND_CODE_ALL.Visible = (optFilter.Value = "B")
    End Sub

    Private Sub grdSATCSANX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSATCSANX.BeforeCellUpdate

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        If COLUMN_NAME.StartsWith("C") And COLUMN_NAME.Length = 4 And Mid(COLUMN_NAME, 2, 4) >= "001" And Mid(COLUMN_NAME, 2, 4) < "999" Then
            If e.Cell.Text <> "" Then
                If e.Cell.Text <> "O" And e.Cell.Text <> "C" Then
                    e.Cancel = True
                End If
            End If
        End If

    End Sub

    Private Sub grdSATCSANX_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSATCSANX.BeforeExitEditMode
        With grdSATCSANX.ActiveCell
            Dim COLUMN_NAME As String = .Column.Key
            If COLUMN_NAME.StartsWith("C") And COLUMN_NAME.Length = 4 And Mid(COLUMN_NAME, 2, 4) >= "001" And Mid(COLUMN_NAME, 2, 4) < "999" Then
                If .EditorResolved.Value & "" <> "" Then
                    .EditorResolved.Value = CStr(.EditorResolved.Value & "").ToUpper
                End If
            End If
        End With
    End Sub


    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        load_handled = True

        For Each rowSATCSANX As DataRow In dst.Tables("SATCSANX").Select("")
            rowSATCSANX.Item("CUST_STORE_CUST_RANK_W") = DBNull.Value
            rowSATCSANX.Item("CUST_STORE_CUST_RANK_M") = DBNull.Value
        Next

        For Each row As DataRow In F.dt.Select("")
            Dim CUST_CODE As String = row.Item(0) & ""
            Dim CUST_STORE_NO As String = row.Item(1) & ""
            Dim CUST_STORE_CUST_RANK_W As String = row.Item(2) & ""
            Dim CUST_STORE_CUST_RANK_M As String = row.Item(3) & ""

            Dim rowSATCSANX As DataRow = dst.Tables("SATCSANX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowSATCSANX IsNot Nothing Then
                rowSATCSANX.Item("CUST_STORE_CUST_RANK_W") = CUST_STORE_CUST_RANK_W
                rowSATCSANX.Item("CUST_STORE_CUST_RANK_M") = CUST_STORE_CUST_RANK_M
            End If
        Next

    End Sub

    Sub Create_WorkFile()

        ASCMAIN1.sql = "SELECT X.*, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE FROM ICTITEM1,ICTCOLL1,(" _
            & "Select RSTRETL1.ITEM_CODE, MIN (RSTRETL1.OPS_YYYYWW) MINYW, MAX (RSTRETL1.OPS_YYYYWW) MAXYW" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' AND NVL(ARTCUST2.CUST_DC_IND,'0') = '1' THEN RSTRETL1.QTY_EOW ELSE 0 END) OHDC" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' AND NVL(ARTCUST2.CUST_DC_IND,'0') = '0' THEN RSTRETL1.QTY_EOW ELSE 0 END) OHST" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW BETWEEN '201531' AND '201604' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_L26" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_TY" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201504' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_LY" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'COMP' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_TY_COMP" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201504' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'COMP' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_LY_COMP" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'DOTCOM' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_TY_DOTCOM" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201504' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'DOTCOM' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_LY_DOTCOM" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201604' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'NEW' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_TY_NEW" _
            & ", SUM (CASE WHEN RSTRETL1.OPS_YYYYWW = '201504' AND ARTCUST2.CUST_STORE_CLASS_CODE = 'NEW' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SLS_LY_NEW" _
            & " from RSTRETL1,ARTCUST2" _
            & " where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" _
            & "   and ARTCUST2.CUST_CODE = RSTRETL1.CUST_CODE" _
            & "   and ARTCUST2.CUST_STORE_NO = RSTRETL1.CUST_STORE_NO" _
            & "   and RSTRETL1.OPS_YYYYWW >= '201501' and RSTRETL1.OPS_YYYYWW <= '" & ryw & "'" _
            & " group by RSTRETL1.ITEM_CODE) X" _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"

    End Sub

End Class