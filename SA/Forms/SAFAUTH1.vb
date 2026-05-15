Public Class SAFAUTH1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim BRAND_CODE As String
    Dim rowICTBRAN1 As DataRow
    Dim sqlSATAUTHX As String
    Dim HC_CODEs As New Dictionary(Of String, Integer)
    Dim SATAUTH2 As String
    Dim i_to_HC_CODE As New Dictionary(Of Integer, String)

    Dim appearance_opened As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGreen}
    Dim appearance_closed As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.DarkOrange}

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SAFAUTHI" Then
            InquiryMode = True
        End If

        ASCMAIN1.sql = "Select * from SATAUTH2 where ROWNUM < 1"
        SATAUTH2 = ASCMAIN1.Temp_Table()

        With dst
            ASCMAIN1.sql = "Select * from SATAUTH1"
            Create_TDA(.Tables.Add, "SATAUTH1", "*")

            ASCMAIN1.sql = "Select * from SATAUTH2 where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATAUTH2", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from ARTCUST2"
            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, , , , "CUST_STORE_CUST_RANK_W,CUST_STORE_CUST_RANK_M")

            sqlSATAUTHX = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.SELL_CODE, SOTSELL1.SELL_NAME, ARTCUST2.MALL_CODE" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CUST_RANK_W, ARTCUST2.CUST_STORE_CUST_RANK_M" & vbCrLf _
                & ", ARTCUST1.SREP_CODE, ARTCUST2.CUST_STORE_STATE" & vbCrLf _
                & " from ARTCUST1, ARTCUST2, SOTSELL1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
                & "   And SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE"
            ASCMAIN1.sql = sqlSATAUTHX & " And ARTCUST2.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SATAUTHX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select * from ICTBRAN1"
            Create_TDA(.Tables.Add, "ICTBRAN1", "**", 0, False)
            .Tables("ICTBRAN1").Columns.Add("SEL")
            .Tables("ICTBRAN1").Columns("SEL").DefaultValue = "0"
            Fill_Records("ICTBRAN1")

            ASCMAIN1.sql = "Select * from ICTCOLL0"
            Create_TDA(.Tables.Add, "ICTCOLL0", "**", 0, False)
            Fill_Records("ICTCOLL0")


            ASCMAIN1.sql = "Select SE.*,C1.HC_CODE from SATEXCL1 SE JOIN ICTCOLL1 C1 ON SE.COLLECTION_CODE=C1.COLLECTION_CODE"
            Create_TDA(.Tables.Add, "SATEXCL1", "**", 0, False)
            Fill_Records("SATEXCL1")

            Dim I As Integer = 0
            For Each rowICTCOLL0 As DataRow In dst.Tables("ICTCOLL0").Select("", "BRAND_CODE,HC_CODE")
                I += 1
                Dim DC As DataColumn = .Tables("SATAUTHX").Columns.Add("C" & Format(I, "000"))
                DC.Caption = rowICTCOLL0.Item("HC_CODE")
                HC_CODEs.Add(rowICTCOLL0.Item("HC_CODE"), I)
                i_to_HC_CODE.Add(I, rowICTCOLL0.Item("HC_CODE"))
            Next
        End With

        grdICTBRAN1.DataSource = dst.Tables("ICTBRAN1")
        Sort_grdColumns(grdICTBRAN1, "BRAND_CODE")

        lblBRAND_CODE.Top = lblCUST_CODE.Top
        txtBRAND_CODE.Top = txtCUST_CODE.Top
        txtBRAND_NAME.Top = txtCUST_NAME.Top


        grdSATAUTHX.DataSource = dst.Tables("SATAUTHX")
        grdSATAUTH2.DataSource = dst.Tables("SATAUTH2")
        Show_Filter(grdSATAUTH2, True)

        With grdSATAUTHX.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add()
            G.Header.Caption = "Customer / Store Information"
            G.Header.Appearance.BackColor2 = Drawing.Color.Aqua
            G.Header.Appearance.BackColor = Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME In New String() {"CUST_CODE", "CUST_STORE_NO", "SREP_CODE", "SELL_CODE", "SELL_NAME", "MALL_CODE", "CUST_STORE_CUST_RANK_W", "CUST_STORE_CUST_RANK_M", "CUST_STORE_STATE"}
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
            With .Columns("SELL_NAME")
                .Header.Caption = "S-Thru"
                .Width = 90
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
            With .Columns("CUST_STORE_STATE")
                .Hidden = True
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

                        Dim hasExclusions = dst.Tables("SATEXCL1").Select($"HC_CODE = '{rowICTCOLL0.Item("HC_CODE")}'").Count > 0
                        If hasExclusions Then
                            .Header.Appearance.BackColor = Drawing.Color.PaleVioletRed
                            .Header.Appearance.ForeColor = Drawing.Color.White
                        End If
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

                        Create_Summary(grdSATAUTHX, COLUMN_NAME, "Custom")
                    End With
                Next
            Next
            .Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        End With

        Create_Summary(grdSATAUTHX, "CUST_CODE", "Count")

        With grdSATAUTHX.DisplayLayout.Bands(0)
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


                For Each COLUMN_NAME As String In New String() {"SREP_CODE", "SELL_CODE", "SELL_NAME", "MALL_CODE"}
                    grdSATAUTHX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                Next
                Excel_Import(grdSATAUTHX)
                For Each COLUMN_NAME As String In New String() {"SREP_CODE", "SELL_CODE", "SELL_NAME", "MALL_CODE"}
                    grdSATAUTHX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
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

        tabAuth.Visible = ScreenMode

        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                grdSATAUTHX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSATAUTHX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSATAUTHX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdSATAUTHX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSATAUTHX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSATAUTHX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl.Panel1Collapsed = ScreenMode
        splSATAUTHX.Visible = ScreenMode

        splSATAUTHX.Panel2Collapsed = True

        If ScreenMode Then
            Set_Groups()

            Dim W As Integer = 30
            If EntryMode = "V" Then W = 65
            For Each HC_CODE As String In HC_CODEs.Keys
                Dim C As Integer = HC_CODEs(HC_CODE)
                grdSATAUTHX.DisplayLayout.Bands(0).Columns("C" & Format(C, "000")).Width = W
            Next
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SATAUTHX", "ARTCUST2", "SATAUTH2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        cmbCopy.Value = ""
        cmbTo.Value = ""

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If optFilter.Value = "B" Then
            ASCMAIN1.sql = sqlSATAUTHX
            Fill_Records("SATAUTHX", "", , ASCMAIN1.sql)
            grdSATAUTHX.Text = "Store / Brand Matrix for " _
                & IIf(BRAND_CODE = "", "All Brands", BRAND_CODE) _
                & ", for All Customers"

            ASCMAIN1.sql = "Select * from SATAUTH1"
            If BRAND_CODE <> "" Then
                ASCMAIN1.sql &= " where HC_CODE in (Select HC_CODE from ICTCOLL0 where BRAND_CODE = '" & BRAND_CODE & "')"
            End If
            Fill_Records("SATAUTH1", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SATAUTH2"
            If BRAND_CODE <> "" Then
                ASCMAIN1.sql &= " where HC_CODE in (Select HC_CODE from ICTCOLL0 where BRAND_CODE = '" & BRAND_CODE & "')"
            End If
            Fill_Records("SATAUTH2", , , ASCMAIN1.sql)

        Else
            Fill_Records("SATAUTHX", New String() {CUST_CODE})
            grdSATAUTHX.Text = "Store / Brand Matrix for Customer " & CUST_CODE

            ASCMAIN1.sql = "Select * from SATAUTH1 where CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SATAUTH1", "", , ASCMAIN1.sql)

            Fill_Records("ARTCUST2", CUST_CODE)
            Fill_Records("SATAUTH2", CUST_CODE)
        End If

        Sort_grdColumns(grdSATAUTH2, "INIT_DATE".ToLower & ",CUST_CODE, CUST_STORE_NO")
        Dim BRAND_CODEs_In As New List(Of String)
        Dim HC_CODEs_In As New List(Of String)

        grdSATAUTHX.SuspendSummaryUpdates()
        grdSATAUTHX.SuspendLayout()

        For Each row As DataRow In dst.Tables("SATAUTH1").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim HC_CODE As String = row.Item("HC_CODE")
            Dim i As Integer = HC_CODEs(HC_CODE)
            Dim rowSATAUTHX As DataRow = dst.Tables("SATAUTHX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If row.Item("OPS_YYYYPP_OPENED") & "" <> "" Then
                If row.Item("OPS_YYYYPP_CLOSED") & "" <> "" Then
                    Dim YYYYMM As String = row.Item("OPS_YYYYPP_CLOSED")
                    Dim MMYY As String = Mid(YYYYMM, 5, 2) & "/" & Mid(YYYYMM, 3, 2)
                    rowSATAUTHX.Item("C" & Format(i, "000")) = "C" & IIf(EntryMode = "V", ":" & MMYY, "")
                    If Not HC_CODEs_In.Contains(HC_CODE) Then HC_CODEs_In.Add(HC_CODE)
                Else
                    Dim YYYYMM As String = row.Item("OPS_YYYYPP_OPENED")
                    Dim MMYY As String = Mid(YYYYMM, 5, 2) & "/" & Mid(YYYYMM, 3, 2)
                    rowSATAUTHX.Item("C" & Format(i, "000")) = "O" & IIf(EntryMode = "V", ":" & MMYY, "")
                    If Not HC_CODEs_In.Contains(HC_CODE) Then HC_CODEs_In.Add(HC_CODE)
                End If
            End If
        Next

        For Each grdCol As UltraWinGrid.UltraGridColumn In grdSATAUTHX.DisplayLayout.Bands(0).Columns
            If HC_CODEs.ContainsKey(grdCol.Header.Caption) Then
                Dim filterString As String = $"HC_CODE = '{grdCol.Header.Caption}'"

                If optFilter.Value <> "B" Then 'By Customer
                    filterString &= $" AND CUST_CODE = '{CUST_CODE}'"
                End If

                Dim hasExclusions As Boolean = dst.Tables("SATEXCL1").Select(filterString).Count > 0
                If hasExclusions Then
                    grdCol.Header.Appearance.BackColor = Drawing.Color.PaleVioletRed
                    grdCol.Header.Appearance.ForeColor = Drawing.Color.White
                Else
                    grdCol.Header.Appearance.BackColor = Drawing.SystemColors.Control
                    grdCol.Header.Appearance.ForeColor = Drawing.Color.Black
                End If
            End If
        Next






        grdSATAUTHX.ResumeSummaryUpdates(True)
        grdSATAUTHX.ResumeLayout()

        For Each HC_CODE As String In HC_CODEs_In
            Dim rowICTCOLL0 As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
            Dim BRAND_CODE As String = rowICTCOLL0.Item("BRAND_CODE") & ""
            If BRAND_CODE <> "" AndAlso Not BRAND_CODEs_In.Contains(BRAND_CODE) Then
                BRAND_CODEs_In.Add(BRAND_CODE)
            End If
        Next

        Sort_grdColumns(grdSATAUTHX, "CUST_CODE,CUST_STORE_NO")

        For Each rowICTBRAN1 As DataRow In dst.Tables("ICTBRAN1").Select("")
            Dim BRAND_CODE As String = rowICTBRAN1.Item("BRAND_CODE")
            If Me.BRAND_CODE <> "" Then
                rowICTBRAN1.Item("SEL") = IIf(BRAND_CODE = Me.BRAND_CODE, "1", "0")
            Else
                rowICTBRAN1.Item("SEL") = IIf(rowICTBRAN1.Item("BRAND_STATUS") & "" = "I", "0", IIf(optFilter.Value = "B" Or BRAND_CODEs_In.Contains(BRAND_CODE), "1", "0"))
            End If
        Next

        dst.Tables("SATAUTHX").AcceptChanges()

        Load_HCs()

        EnforceConstraints(True)
        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-composing Authorization Records")

        Dim sql1cols As String = "CUST_CODE,CUST_STORE_NO,HC_CODE,OPS_YYYYPP_OPENED,OPS_YYYYPP_CLOSED"
        ASCDATA1.ExecuteSQL("Delete from " & SATAUTH2)
        Dim sql1 As String = "Select * from SATAUTH1"
        If optFilter.Value = "C" Then
            sql1 &= " where CUST_CODE = '" & CUST_CODE & "'"
        Else
            If BRAND_CODE <> "" Then
                sql1 &= " where HC_CODE in (Select HC_CODE from ICTCOLL0 where BRAND_CODE = '" & BRAND_CODE & "')"
            End If
        End If
        ASCDATA1.ExecuteSQL("Insert into " & SATAUTH2 & " (" & sql1cols & ") " & sql1)

        Dim YP As String = cbeYP.Value

        For Each row As DataRow In dst.Tables("SATAUTHX").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")

            If optFilter.Value = "C" Then
                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                If rowARTCUST2.Item("CUST_STORE_CUST_RANK_W") & "" <> row.Item("CUST_STORE_CUST_RANK_W") & "" Then rowARTCUST2.Item("CUST_STORE_CUST_RANK_W") = row.Item("CUST_STORE_CUST_RANK_W")
                If rowARTCUST2.Item("CUST_STORE_CUST_RANK_M") & "" <> row.Item("CUST_STORE_CUST_RANK_M") & "" Then rowARTCUST2.Item("CUST_STORE_CUST_RANK_M") = row.Item("CUST_STORE_CUST_RANK_M")
            End If

            For Each HC_CODE As String In HC_CODEs.Keys

                Dim i As Integer = HC_CODEs(HC_CODE)
                Dim C000 As String = "C" & Format(i, "000")

                Dim rowSATAUTH1 As DataRow = dst.Tables("SATAUTH1").Rows.Find _
                            (New String() {CUST_CODE, CUST_STORE_NO, HC_CODE})

                If row.Item(C000) & "" <> "" Then
                    If rowSATAUTH1 Is Nothing Then
                        rowSATAUTH1 = dst.Tables("SATAUTH1").NewRow
                        rowSATAUTH1.Item("CUST_CODE") = CUST_CODE
                        rowSATAUTH1.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSATAUTH1.Item("HC_CODE") = HC_CODE
                        dst.Tables("SATAUTH1").Rows.Add(rowSATAUTH1)
                    End If
                End If

                If row.Item(C000) & "" = "O" Then
                    If rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") & "" <> "" Then
                        rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") = ""
                        rowSATAUTH1.Item("OPS_YYYYPP_OPENED") = YP ' ASCMAIN1.CYP
                    ElseIf rowSATAUTH1.Item("OPS_YYYYPP_OPENED") & "" = "" Then
                        rowSATAUTH1.Item("OPS_YYYYPP_OPENED") = YP ' ASCMAIN1.CYP
                    End If
                ElseIf row.Item(C000) & "" = "C" Then
                    If rowSATAUTH1.Item("OPS_YYYYPP_OPENED") & "" <> "" _
                        And rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") & "" = "" Then
                        rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") = YP ' ASCMAIN1.CYP
                    End If
                ElseIf rowSATAUTH1 IsNot Nothing AndAlso row.Item(C000) & "" = "" Then
                    If rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") & "" <> "" _
                    Or rowSATAUTH1.Item("OPS_YYYYPP_OPENED") & "" <> "" Then
                        rowSATAUTH1.Item("OPS_YYYYPP_OPENED") = DBNull.Value
                        rowSATAUTH1.Item("OPS_YYYYPP_CLOSED") = DBNull.Value
                    End If
                End If
            Next
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Try
            BeginTrans()
            Update_Record_TDA("SATAUTH1")
            If optFilter.Value = "C" Then Update_Record_TDA("ARTCUST2")

            ASCMAIN1.sql = "Insert into SATAUTH2 Select " _
                & sql1cols & ", '" & ASCMAIN1.USER_ID & "' INIT_OPER, SYSDATE INIT_DATE,'" & XNO & "' XNO" _
                & " from (" & sql1 & " minus Select " & sql1cols & " from " & SATAUTH2 & ")"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSQL("Delete from SATAUTH1 where (CUST_CODE,CUST_STORE_NO,HC_CODE) in (Select CUST_CODE,CUST_STORE_NO,HC_CODE from (" & sql1 & ")) AND OPS_YYYYPP_OPENED IS NULL AND OPS_YYYYPP_CLOSED IS NULL")

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATAUTHX, "SSBBBBBB", "Show Filter", "Show GroupBox", _
                        "Set Copy From", "Set Copy To", "Restore", _
                        "Open All Stores for Collection", "Close All Stores for Collection", "Close All Collections for Store")
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
                Case "grdSATAUTHX"

                    tlb_btn = tlb_pop.Tools("Set Copy From")
                    If grd.ActiveCell Is Nothing Or (EntryMode <> "E") Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim HC_CODE = grd.ActiveCell.Column.Header.Caption
                        tlb_btn.SharedProps.Caption = "Set Copy From to " & HC_CODE
                        tlb_btn.Tag = HC_CODE
                        tlb_btn.SharedProps.Visible = True
                    End If


                    tlb_btn = tlb_pop.Tools("Set Copy To")
                    If grd.ActiveCell Is Nothing Or (EntryMode <> "E") Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim HC_CODE = grd.ActiveCell.Column.Header.Caption
                        tlb_btn.SharedProps.Caption = "Set Copy To to " & HC_CODE
                        tlb_btn.Tag = HC_CODE
                        tlb_btn.SharedProps.Visible = True
                    End If

                    tlb_btn = tlb_pop.Tools("Restore")
                    If grd.ActiveCell Is Nothing Or EntryMode <> "E" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length = 4 And C.StartsWith("C") And Mid(C, 2, 3) >= "001" And Mid(C, 2, 3) <= "999" Then
                            tlb_btn.SharedProps.Visible = True
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If

                    tlb_btn = tlb_pop.Tools("Open All Stores for Collection")
                    If grd.ActiveCell Is Nothing Or EntryMode <> "E" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length = 4 And C.StartsWith("C") And Mid(C, 2, 3) >= "001" And Mid(C, 2, 3) <= "999" Then
                            tlb_btn.SharedProps.Visible = True
                            Dim HC_CODE As String = i_to_HC_CODE(Val(Mid(C, 2, 3)))
                            tlb_btn.Tag = HC_CODE
                            tlb_btn.SharedProps.Caption = "Open All Stores for " & HC_CODE
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If

                    tlb_btn = tlb_pop.Tools("Close All Stores for Collection")
                    If grd.ActiveCell Is Nothing Or (EntryMode <> "E") Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim C As String = grd.ActiveCell.Column.Key
                        If C.Length = 4 And C.StartsWith("C") And Mid(C, 2, 3) >= "001" And Mid(C, 2, 3) <= "999" Then
                            tlb_btn.SharedProps.Visible = True
                            Dim HC_CODE As String = i_to_HC_CODE(Val(Mid(C, 2, 3)))
                            tlb_btn.Tag = HC_CODE
                            tlb_btn.SharedProps.Caption = "Close All Stores for " & HC_CODE
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    End If

                    tlb_btn = tlb_pop.Tools("Close All Collections for Store")
                    If grd.ActiveRow Is Nothing Or (EntryMode <> "E") Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                    End If
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
            Case "Set Copy From"
                Dim HC_CODE As String = e.Tool.Tag
                cmbCopy.Value = HC_CODE

            Case "Set Copy To"
                Dim HC_CODE As String = e.Tool.Tag
                cmbTo.Value = HC_CODE

            Case "Restore"
                If grd.ActiveCell IsNot Nothing Then
                    Dim C As String = grd.ActiveCell.Column.Key
                    '  Dim rowSATAUTH1 As DataRow = dst.Tables ("SATAUTH1").Rows.Find(New String () {"",""})
                    Dim R As Integer = grd.Rows.IndexOf(grd.ActiveRow)

                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                    Dim CUST_STORE_NO As String = grd.ActiveRow.Cells("CUST_STORE_NO").Value
                    Dim rowSATAUTHX As DataRow = dst.Tables("SATAUTHX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                    grd.ActiveCell.Value = rowSATAUTHX.Item(C, DataRowVersion.Original)
                End If

            Case "Open All Stores for Collection", "Close All Stores for Collection"
                If MsgBox("Are you sure that you want to " & vbCrLf & e.Tool.SharedProps.Caption & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Updating")

                Dim OC As String = Mid(e.Tool.Key, 1, 1)
                Dim tlb_btn As UltraWinToolbars.ButtonTool = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim HC_CODE As String = tlb_btn.Tag
                Dim C As String = "C" & Format(HC_CODEs(HC_CODE), "000")
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If OC = "O" Then
                        If grow.Cells(C).Value & "" = "" Then
                            grow.Cells(C).Value = "O"
                            grow.Update()
                        End If
                    Else
                        If grow.Cells(C).Value & "" = "O" Then
                            grow.Cells(C).Value = "C"
                            grow.Update()
                        End If
                    End If
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")


            Case "Close All Collections for Store"
                Dim CUST_STORE_NO As String = grd.ActiveRow.Cells("CUST_STORE_NO").Value
                If MsgBox("Are you sure that you want to " & vbCrLf & e.Tool.SharedProps.Caption & " " & CUST_STORE_NO & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Updating")

                For i As Integer = 1 To HC_CODEs.Count
                    Dim C As String = "C" & Format(i, "000")
                    If grd.ActiveRow.Cells(C).Value & "" = "O" Then
                        grd.ActiveRow.Cells(C).Value = "C"
                        grd.ActiveRow.Update()
                    End If
                Next


                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

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
            grdSATAUTHX.DisplayLayout.Bands(0).Groups(BRAND_CODE).Hidden = (rowICTBRAN1.Item("SEL") & "" <> "1")
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSATAUTHX_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs) Handles grdSATAUTHX.ClickCell
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

    Private Sub grdSATAUTHX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSATAUTHX.InitializeRow
        For Each HC_CODE As String In HC_CODEs.Keys
            Dim I As Integer = HC_CODEs(HC_CODE)
            Dim cellId As String = "C" & Format(I, "000")
            If CStr(e.Row.Cells(cellId).Value & "").StartsWith("O") Then
                e.Row.Cells(cellId).Appearance = appearance_opened
            ElseIf CStr(e.Row.Cells(cellId).Value & "").StartsWith("C") Then
                e.Row.Cells(cellId).Appearance = appearance_closed
                'If e.Row.Cells("CUST_STORE_NO").Value = "000006" Then
                '    e.Row.Cells("C" & Format(I, "000")).Appearance.ForeColor = Drawing.Color.Red
                'End If

            Else
                e.Row.Cells("C" & Format(I, "000")).Appearance = Nothing
            End If

            'Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value
            'Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value
            'Dim STATE_CODE As String = e.Row.Cells("CUST_STORE_STATE").Value & ""
            'Dim hcExclusions As DataRow() = dst.Tables("SATEXCL1").Select($"HC_CODE='{HC_CODE}' AND STATE_CODE='{STATE_CODE}'")
            'If hcExclusions.Count > 0 Then
            '    e.Row.Cells(cellId).Appearance.ForeColor = Drawing.Color.Red
            'End If
        Next
    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdSATAUTHX"
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
            Case "grdSATAUTHX"
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

    Private Sub grdSATAUTHX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSATAUTHX.BeforeCellUpdate

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        If COLUMN_NAME.StartsWith("C") And COLUMN_NAME.Length = 4 And Mid(COLUMN_NAME, 2, 4) >= "001" And Mid(COLUMN_NAME, 2, 4) < "999" Then
            If e.Cell.Text <> "" Then
                If e.Cell.Text <> "O" And e.Cell.Text <> "C" Then
                    e.Cancel = True
                End If
            End If
        End If
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '        Dim ITEM_CODE As String = ""
        '        If ITEM_CODE = "" Then
        '            e.Cancel = True
        '        End If
        'End Select
    End Sub

    Private Sub grdSATAUTHX_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSATAUTHX.BeforeExitEditMode
        With grdSATAUTHX.ActiveCell
            Dim COLUMN_NAME As String = .Column.Key
            If COLUMN_NAME.StartsWith("C") And COLUMN_NAME.Length = 4 And Mid(COLUMN_NAME, 2, 4) >= "001" And Mid(COLUMN_NAME, 2, 4) < "999" Then
                If .EditorResolved.Value & "" <> "" Then
                    .EditorResolved.Value = CStr(.EditorResolved.Value & "").ToUpper
                End If
            End If
        End With
    End Sub

    Private Sub cmdPerformCopy_Click(sender As Object, e As EventArgs) Handles cmdPerformCopy.Click
        Dim HC_CODE_copy As String = cmbCopy.Value & ""
        Dim HC_CODE_to As String = cmbTo.Value & ""

        If HC_CODE_copy = "" Or HC_CODE_to = "" Or (HC_CODE_copy = HC_CODE_to) Then
            MsgBox("You must specify a High Collection to Copy From," & vbCrLf & " as well as a (different) High Collection to Copy To", MsgBoxStyle.OkOnly, "Cannot Perform Copy")
            Exit Sub
        End If

        If MsgBox("This option will establish Authorizations for " & HC_CODE_to & " in the Current Period" _
                  & vbCrLf & " for all Stores marked as Open in " & HC_CODE_copy & "." _
                  & vbCrLf & vbCrLf & "Continue with this Copy?", _
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If
        Dim HCI_copy As Integer = HC_CODEs(HC_CODE_copy)
        Dim HCI_to As Integer = HC_CODEs(HC_CODE_to)

        Dim C_copy As String = "C" & Format(HCI_copy, "000")
        Dim C_to As String = "C" & Format(HCI_to, "000")

        For Each row As DataRow In dst.Tables("SATAUTHX").Select(C_to & " IS NOT NULL")
            row.Item(C_to) = DBNull.Value
        Next

        ASCDATA1.DeleteRows("SATAUTH1", "HC_CODE = '" & HC_CODE_to & "'")

        For Each row As DataRow In dst.Tables("SATAUTHX").Select(C_copy & " = 'O'")
            row.Item(C_to) = "O"
        Next

    End Sub

    Sub Load_HCs()

        Dim HCs As New List(Of String)
        For Each row As DataRow In dst.Tables("ICTCOLL0") _
            .Select(IIf(optFilter.Value = "B" And BRAND_CODE <> "", "BRAND_CODE = '" & BRAND_CODE & "'", ""), "HC_CODE")
            Dim HC_CODE As String = row.Item("HC_CODE")
            HCs.Add(HC_CODE)
        Next
        cmbCopy.DataSource = HCs
        cmbTo.DataSource = HCs
    End Sub
    Private Sub cmbCopy_BeforeDropDown(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cmbCopy.BeforeDropDown

    End Sub

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        load_handled = True

        For Each rowSATAUTHX As DataRow In dst.Tables("SATAUTHX").Select("")
            rowSATAUTHX.Item("CUST_STORE_CUST_RANK_W") = DBNull.Value
            rowSATAUTHX.Item("CUST_STORE_CUST_RANK_M") = DBNull.Value
        Next

        For Each row As DataRow In F.dt.Select("")
            Dim CUST_CODE As String = row.Item(0) & ""
            Dim CUST_STORE_NO As String = row.Item(1) & ""
            Dim CUST_STORE_CUST_RANK_W As String = row.Item(2) & ""
            Dim CUST_STORE_CUST_RANK_M As String = row.Item(3) & ""

            Dim rowSATAUTHX As DataRow = dst.Tables("SATAUTHX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowSATAUTHX IsNot Nothing Then
                rowSATAUTHX.Item("CUST_STORE_CUST_RANK_W") = CUST_STORE_CUST_RANK_W
                rowSATAUTHX.Item("CUST_STORE_CUST_RANK_M") = CUST_STORE_CUST_RANK_M
            End If
        Next

    End Sub

End Class