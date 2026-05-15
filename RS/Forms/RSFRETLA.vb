Public Class RSFRETLA
    Dim RSTRETLA As String
    Dim OPS_YYYY As String
    Dim Imax As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_RSTRETLA("0000")

            ASCMAIN1.sql = "Select ICTRETLA.* from ICTRETLA " _
            & " where OPS_YYYYPP like :PARM1 and ITEM_CODE in (Select ITEM_CODE from " & RSTRETLA & ")"
            Create_TDA(.Tables.Add, "ICTRETLA", "**", 0, True, "V")

            ASCMAIN1.sql = "Select RSTRETLA.*, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_RETAIL_PRICE" _
            & " from ICTITEM1," & RSTRETLA & " RSTRETLA" _
            & " where ICTITEM1.ITEM_CODE = RSTRETLA.ITEM_CODE"
            Create_TDA(.Tables.Add, "RSTRETLA", "**", 0, False, "", 1)
            For I As Integer = 1 To 12
                .Tables("RSTRETLA").Columns.Add("HST_" & Format(I, "00"), GetType(System.Decimal))
                Dim f As String = "IIF(ISNULL(QTY_00,0)=0 OR ISNULL(AMT_00,0)=0,HST_00,AMT_00/QTY_00)"
                .Tables("RSTRETLA").Columns.Add("IMP_" & Format(I, "00"), GetType(System.Decimal), Replace(f, "_00", "_" & Format(I, "00")))
                f = "ISNULL(QTY_00,0) * ISNULL(HST_00,0)"
                .Tables("RSTRETLA").Columns.Add("EXT_" & Format(I, "00"), GetType(System.Decimal), Replace(f, "_00", "_" & Format(I, "00")))
            Next
            .Tables("RSTRETLA").Columns.Add("UPDATE")

            ASCMAIN1.sql = "Select RSTRETL1.*" _
            & " from RSTRETL1" _
            & " where RSTRETL1.ITEM_CODE = :PARM1 and RSTRETL1.OPS_YYYYPP = :PARM2" _
            & " and (QTY_SOLD <> 0 OR AMT_SOLD <> 0)"
            Create_TDA(.Tables.Add, "RSTRETL1", "**", 0, False, "VV")

        End With

        grdRSTRETLA.DataSource = dst.Tables("RSTRETLA")
        grdRSTRETL1.DataSource = dst.Tables("RSTRETL1")

        Create_Summary(grdRSTRETLA, "ITEM_CODE", "Count")

        Create_Summary(grdRSTRETL1, "QTY_SOLD")
        Create_Summary(grdRSTRETL1, "AMT_SOLD")

        grdRSTRETLA.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdRSTRETLA.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

        With grdRSTRETLA.DisplayLayout.Bands("RSTRETLA")

            '.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            '.Override.AllowDelete = DefaultableBoolean.False

            Dim G As UltraWinGrid.UltraGridGroup

            G = .Groups.Add
            G.Header.Fixed = True
            G.Header.Caption = "Item"
            G.Header.Appearance.BackColor2 = Color.Orange
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME As String In New String() _
                    {"ITEM_CODE", "UPDATE"}
                .Columns(COLUMN_NAME).Group = G
                If COLUMN_NAME = "UPDATE" Then
                Else
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next

            G = .Groups.Add
            G.Header.Caption = "Item Master Data"
            G.Header.Appearance.BackColor2 = Color.LightGreen
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            For Each COLUMN_NAME As String In New String() _
                    {"ITEM_DESC", "COLLECTION_CODE", "ITEM_RETAIL_PRICE"}
                .Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next

            For I As Integer = 1 To 12
                G = .Groups.Add
                G.Key = Format(I, "00")
                G.Header.Appearance.BackColor2 = Color.LightBlue
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                'Dim chk As New UltraWinEditors.UltraCheckEditor
                'chk.Text = "Test"
                'chk.Name = "chk" & G.Key
                'chk.Appearance.BackColor = Color.Transparent
                'chk.Left = I * 100
                'chk.Parent = grdRSTRETLA
                'chk.Visible = True

                For Each COLUMN_NAME_PFX As String In New String() _
                        {"QTY", "AMT", "IMP", "HST", "EXT"}
                    Dim COLUMN_NAME As String = COLUMN_NAME_PFX & "_" & Format(I, "00")

                    .Columns(COLUMN_NAME).Header.Caption = ASCMAIN1.Make_Caption(COLUMN_NAME_PFX)
                    .Columns(COLUMN_NAME).Width = 70
                    If COLUMN_NAME_PFX = "QTY" Or COLUMN_NAME_PFX = "AMT" Or COLUMN_NAME_PFX = "EXT" Then
                        .Columns(COLUMN_NAME).Format = "###,##0"
                    Else
                        .Columns(COLUMN_NAME).Format = "###,##0.00"
                    End If
                    .Columns(COLUMN_NAME).Group = G
                    If COLUMN_NAME_PFX = "HST" Then
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                    End If
                    If COLUMN_NAME_PFX = "QTY" Or COLUMN_NAME_PFX = "AMT" Or COLUMN_NAME_PFX = "EXT" Then
                        Create_Summary(grdRSTRETLA, COLUMN_NAME)
                    End If
                Next
            Next
        End With

        Dim YEARs As New List(Of String)
        For Y As Integer = 0 To 2
            YEARs.Add(Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) - Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs 
        'spl.Panel1Collapsed = True

        DisplayHeaderCheckBox(grdRSTRETLA, New String() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        'For Each de As String In aCheckBoxOnGroupHeader_CreationFilter.GroupKeys.Keys
        '    aCheckBoxOnGroupHeader_CreationFilter.GroupKeys(de) = "1"
        'Next

        For I As Integer = 1 To 12
            aCheckBoxOnGroupHeader_CreationFilter.GroupKeys(Format(I, "00")) = "1"
        Next

        splRTL.Panel2Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Value
                If OPS_YYYY = "" Then
                    EMsg &= vbCr & "You Must Select a Valid Year before Loading"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTRETLA", OPS_YYYY) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                If dst.Tables("RSTRETLA").Select("UPDATE = '1'").Length = 0 Then
                    EMsg &= vbCr & "No Items have been Checked Off for Update"
                End If

                Dim M As Integer = 0
                Dim M_TOTAL As Integer = 0
                For I As Integer = 1 To 12
                    If Not grdRSTRETLA.DisplayLayout.Bands(0).Groups(Format(I, "00")).Hidden Then
                        M_TOTAL += 1
                        If aCheckBoxOnGroupHeader_CreationFilter.GroupKeys(Format(I, "00")) = "1" Then
                            M += 1
                        End If
                    End If
                Next
                If M_TOTAL = 0 Or M = 0 Then
                    EMsg &= vbCr & "No Months have been Checked Off for Update"
                Else
                    If M <> M_TOTAL Then
                        If MsgBox("You are updating only " & CStr(M) & " of the " & CStr(M_TOTAL) & " Months on Display" & vbCr & vbCr & "OK to Continue with Update?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

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

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Display Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.cbeFor("OPS_YYYY"), ScreenMode)

        grdRSTRETLA.Visible = ScreenMode
        grdRSTRETLA.Text = "Retail Price and Sales History for " & OPS_YYYY
        splRTL.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("ICTRETLA").Rows.Clear()
        dst.Tables("RSTRETLA").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.cbeFor("OPS_YYYY").Value = Val(Mid(ASCMAIN1.CYP, 1, 4))
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        Create_RSTRETLA(OPS_YYYY)
        Fill_Records("ICTRETLA", OPS_YYYY & "%")
        Fill_Records("RSTRETLA")

        Imax = 12
        If OPS_YYYY = Mid(ASCMAIN1.CYP, 1, 4) Then
            Imax = Val(Mid(ASCMAIN1.CYP, 5, 2))
        End If

        With grdRSTRETLA.DisplayLayout.Bands(0)
            For i As Integer = 1 To 12
                If i > Imax Then
                    .Groups(Format(i, "00")).Hidden = True
                Else
                    .Groups(Format(i, "00")).Hidden = False
                    .Groups(Format(i, "00")).Header.Caption = ASCMAIN1.Get_Legend(OPS_YYYY & Format(i, "00"))
                End If
            Next
        End With

        For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA").Rows
            Dim ITEM_CODE As String = rowRSTRETLA.Item("ITEM_CODE")
            For I As Integer = 1 To Imax
                Dim OPS_YYYYPP As String = OPS_YYYY & Format(I, "00")
                Dim rowICTRETLA As DataRow = dst.Tables("ICTRETLA").Rows.Find(New Object() {ITEM_CODE, OPS_YYYYPP})
                If rowICTRETLA IsNot Nothing Then
                    rowRSTRETLA.Item("HST_" & Format(I, "00")) = rowICTRETLA.Item("ITEM_RETAIL_PRICE")
                End If
            Next
        Next

        dst.Tables("RSTRETLA").AcceptChanges()

        For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA").Rows
            Dim ITEM_RETAIL_PRICE As Decimal = 0
            For I As Integer = 1 To Imax
                If I > 1 Then
                    If Val(rowRSTRETLA.Item("HST_" & Format(I, "00")) & "") = 0 And ITEM_RETAIL_PRICE <> 0 Then
                        rowRSTRETLA.Item("HST_" & Format(I, "00")) = ITEM_RETAIL_PRICE
                    End If
                End If
                ITEM_RETAIL_PRICE = Val(rowRSTRETLA.Item("HST_" & Format(I, "00")) & "")
            Next
        Next

        optDisplay.Value = "A"

        Dim dvw As DataView = DirectCast(grdRSTRETLA.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""

        Sort_grdColumns(grdRSTRETLA, "ITEM_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA").Select("UPDATE = '1'")
            Dim ITEM_CODE As String = rowRSTRETLA.Item("ITEM_CODE")
            For I As Integer = 1 To Imax
                If aCheckBoxOnGroupHeader_CreationFilter.GroupKeys(Format(I, "00")) = "1" Then
                    'Dim price_changed As Boolean = False
                    Dim OPS_YYYYPP As String = OPS_YYYY & Format(I, "00")
                    Dim ITEM_RETAIL_PRICE As Decimal = Val(rowRSTRETLA.Item("HST_" & Format(I, "00")) & "")
                    If ITEM_RETAIL_PRICE <> 0 Then
                        Dim rowICTRETLA As DataRow = dst.Tables("ICTRETLA").Rows.Find(New Object() {ITEM_CODE, OPS_YYYYPP})
                        If rowICTRETLA Is Nothing Then
                            'price_changed = True
                            rowICTRETLA = dst.Tables("ICTRETLA").NewRow
                            rowICTRETLA.Item("ITEM_CODE") = ITEM_CODE
                            rowICTRETLA.Item("OPS_YYYYPP") = OPS_YYYYPP
                            rowICTRETLA.Item("ITEM_RETAIL_PRICE") = ITEM_RETAIL_PRICE
                            dst.Tables("ICTRETLA").Rows.Add(rowICTRETLA)
                        Else
                            If Val(rowICTRETLA.Item("ITEM_RETAIL_PRICE") & "") <> ITEM_RETAIL_PRICE Then
                                'price_changed = True
                                rowICTRETLA.Item("ITEM_RETAIL_PRICE") = ITEM_RETAIL_PRICE
                            End If
                        End If

                        ASCMAIN1.sql = "Update RSTRETL1 Set AMT_SOLD = QTY_SOLD * :PARM1" _
                        & " where ITEM_CODE = :PARM2 and OPS_YYYYPP = :PARM3"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NVV", New Object() {ITEM_RETAIL_PRICE, ITEM_CODE, OPS_YYYYPP})
                    End If
                End If
            Next
        Next

        Update_Record_TDA("ICTRETLA")

        If chkRebuildSummary.Checked Then
            ASCMAIN1.sql = "Delete from RSTRETL4 " _
            & " where OPS_YYYYPP >= '" & OPS_YYYY & "01' and OPS_YYYYPP <= '" & OPS_YYYY & "12'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from RSTRETL2 " _
            & " where OPS_YYYYPP >= '" & OPS_YYYY & "01' and OPS_YYYYPP <= '" & OPS_YYYY & "12'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into RSTRETL4 Select " _
            & "RSTRETL1.CUST_CODE," _
            & "RSTRETL1.CUST_STORE_NO," _
            & "ICTITEM1.COLLECTION_CODE," _
            & "RSTRETL1.OPS_YYYYWW," _
            & "RSTRETL1.OPS_YYYYPP," _
            & "SUM (RSTRETL1.QTY_SOLD) QTY_SOLD," _
            & "SUM (RSTRETL1.AMT_SOLD) AMT_SOLD," _
            & "'JH' BRAND_CODE," _
            & "NULL SELL_CODE," _
            & "NULL SREP_CODE," _
            & "SUM (RSTRETL1.QTY_EOW) QTY_EOW" _
            & " from RSTRETL1,ICTITEM1 WHERE ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" _
            & " and RSTRETL1.OPS_YYYYPP >= '" & OPS_YYYY & "01' and RSTRETL1.OPS_YYYYPP <= '" & OPS_YYYY & "12'" _
            & " group by " _
            & "RSTRETL1.CUST_CODE," _
            & "RSTRETL1.CUST_STORE_NO," _
            & "ICTITEM1.COLLECTION_CODE," _
            & "RSTRETL1.OPS_YYYYWW," _
            & "RSTRETL1.OPS_YYYYPP"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into RSTRETL2 Select " _
            & "OPS_YYYYPP, CUST_CODE, CUST_STORE_NO, COLLECTION_CODE, SUM (AMT_SOLD) FROM RSTRETL4" _
            & " where OPS_YYYYPP >= '" & OPS_YYYY & "01' and OPS_YYYYPP <= '" & OPS_YYYY & "12'" _
            & " group by OPS_YYYYPP, CUST_CODE, CUST_STORE_NO, COLLECTION_CODE"
            ASCDATA1.ExecuteSQL()
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

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTRETLA, "SSBSB", "Show Filter", "Show GroupBox", "Mark Selected for Update", "Show Only if Red", "Show Details")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Only if Red") Then
            Dim dvw As DataView = DirectCast(grdRSTRETLA.DataSource, DataTable).DefaultView
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Only if Red"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (dvw.RowFilter <> "")
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdRSTBUDR1"
                    If grdRSTRETLA.Tag = "" Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Mark Selected for Update"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("UPDATE").Value = "1"
                    grow.Update()
                    'Dim ITEM_CODE As String = grow.Cells("ITEM_CODE").Value
                    'Dim rowRSTRETLA As DataRow = dst.Tables("RSTRETLA").Rows.Find(ITEM_CODE)
                    'rowRSTRETLA.SetModified()
                Next

            Case "Show Only if Red"
                Dim dvw As DataView = DirectCast(grdRSTRETLA.DataSource, DataTable).DefaultView
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim HST_ne_IMP As String = ""
                    For i As Integer = 1 To Imax
                        HST_ne_IMP &= " or HST_" & Format(i, "00") & "<>" & "IMP_" & Format(i, "00")
                    Next
                    dvw.RowFilter = Mid(HST_ne_IMP, 5)
                Else
                    dvw.RowFilter = ""
                End If

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Details"
                If grdRSTRETLA.ActiveCell IsNot Nothing Then
                    Dim GRP As String = grdRSTRETLA.ActiveCell.Column.Group.Key
                    Dim OPS_YYYYPP As String = OPS_YYYY & GRP
                    splRTL.Panel2Collapsed = False
                    Dim ITEM_CODE As String = grdRSTRETLA.ActiveRow.Cells("ITEM_CODE").Value
                    Fill_Records("RSTRETL1", New String() {ITEM_CODE, OPS_YYYYPP})
                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case ""
        End Select

    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case ""
        End Select
    End Sub

    Public Overrides Sub chk_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Dim COLUMN_NAME_PFX As String = Mid(COLUMN_NAME, 6, 3)
        Dim hide_column As Boolean = Not Absx1.chkFor(COLUMN_NAME).Checked
        With grdRSTRETLA.DisplayLayout.Bands(0)
            For i As Integer = 1 To Imax
                .Columns(COLUMN_NAME_PFX & "_" & Format(i, "00")).Hidden = hide_column
            Next
        End With
    End Sub
#End Region

    Sub Create_RSTRETLA(ByVal OPS_YYYY As String)

        Dim sql As String = ""
        For i As Integer = 1 To 12
            sql &= ", SUM (DECODE(OPS_YYYYPP,'" & OPS_YYYY & Format(i, "00") & "',QTY_SOLD)) QTY_" & Format(i, "00")
            sql &= ", SUM (DECODE(OPS_YYYYPP,'" & OPS_YYYY & Format(i, "00") & "',AMT_SOLD)) AMT_" & Format(i, "00")
        Next
        Dim sqlRSTRETLA As String = "Select ITEM_CODE" & sql _
        & " from RSTRETL1 " _
        & " where OPS_YYYYPP >= '" & OPS_YYYY & "01' AND OPS_YYYYPP <= '" & OPS_YYYY & "12'" _
        & " and (QTY_SOLD <> 0 or AMT_SOLD <> 0)" _
        & " group by ITEM_CODE"

        If RSTRETLA = "" Then
            RSTRETLA = ASCMAIN1.Temp_Table(sqlRSTRETLA)
            ASCDATA1.ExecuteSQL("Alter Table " & RSTRETLA & " Add Primary Key (ITEM_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTRETLA)
            ASCDATA1.ExecuteSQL("Insert into " & RSTRETLA & " " & sqlRSTRETLA)
        End If

    End Sub

    Private Sub optDisplay_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDisplay.ValueChanged

        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim dvw As DataView = dst.Tables("RSTRETLA").DefaultView
        Select Case optDisplay.Value
            Case "A"
                dvw.RowFilter = ""
            Case "1"
                Dim sql As String = ""
                For i As Integer = 1 To Imax
                    sql &= " or " & Replace("ISNULL(HST_00,0) <> ISNULL(IMP_00,0)", "_00", "_" & Format(i, "00"))
                Next
                dvw.RowFilter = Mid(sql, 4)
            Case "2"
                Dim sql As String = ""
                For i As Integer = 2 To Imax
                    sql &= " or " & Replace(Replace("ISNULL(HST_00,0) <> ISNULL(HST_01,0)", "_00", "_" & Format(i - 1, "00")), "_01", "_" & Format(i, "00"))
                Next
                dvw.RowFilter = Mid(sql, 4)
            Case "3"
                dvw.RowFilter = "UPDATE = '1'"
        End Select
    End Sub

    Private Sub grdRSTRETLA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLA.AfterCellUpdate
        If e.Cell.Row.IsDataRow Then
            If e.Cell.Column.Key <> "UPDATE" Then
                If chkCopyForward.Checked Then
                    Dim COLUMN_NAME As String = e.Cell.Column.Key
                    Dim C As Integer = Val(Mid(COLUMN_NAME, 5, 2))
                    If C < Imax Then
                        For I As Integer = C + 1 To Imax
                            e.Cell.Row.Cells("HST_" & Format(I, "00")).Value = e.Cell.Value
                        Next
                    End If
                End If
                If e.Cell.Row.Cells("UPDATE").Value & "" <> "1" Then
                    e.Cell.Row.Cells("UPDATE").Value = "1"
                End If
            End If
        End If
    End Sub

    Private Sub grdRSTRETLA_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLA.AfterRowActivate
        splRTL.Panel2Collapsed = True
    End Sub

    Private Sub grdRSTRETLA_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTRETLA.InitializeLayout

    End Sub

    Private Sub grdRSTRETLA_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTRETLA.InitializeRow
        If e.Row.IsDataRow Then
            Dim HST_LAST As Decimal = 0
            For i As Integer = 1 To Imax
                Dim HST As Decimal = Val(e.Row.Cells("HST_" & Format(i, "00")).Value & "")
                Dim IMP As Decimal = Val(e.Row.Cells("IMP_" & Format(i, "00")).Value & "")
                If HST <> IMP Then
                    e.Row.Cells("HST_" & Format(i, "00")).Appearance.ForeColor = Color.Red
                Else
                    e.Row.Cells("HST_" & Format(i, "00")).Appearance.ForeColor = Color.Empty
                End If
                If i > 1 Then
                    If HST <> HST_LAST Then
                        e.Row.Cells("HST_" & Format(i, "00")).Appearance.BackColor = Color.Yellow
                    Else
                        e.Row.Cells("HST_" & Format(i, "00")).Appearance.BackColor = Color.Empty
                    End If
                End If
                HST_LAST = HST
            Next
            If e.Row.Cells("UPDATE").Value & "" = "1" Then
                e.Row.Cells("UPDATE").Appearance.BackColor = Color.Red
            Else
                e.Row.Cells("UPDATE").Appearance.BackColor = Color.Empty
            End If
        End If
    End Sub

    Dim WithEvents aCheckBoxOnGroupHeader_CreationFilter As New CheckBoxOnGroupHeader_CreationFilter()

    ''' <summary>
    ''' This event on the CreationFilter fires when the CheckBox in a Header is clicked. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub aCheckBoxOnHeader_CreationFilter_HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As CheckBoxOnGroupHeader_CreationFilter.HeaderCheckBoxEventArgs) Handles aCheckBoxOnGroupHeader_CreationFilter.HeaderCheckBoxClicked
        ' Check to see if the column is of type boolean.  If it is, set all the cells in that column to
        ' whatever value the header checkbox is.
        'Dim aRow As UltraWinGrid.UltraGridRow
        'Dim level As Integer = e.Header.Column.Level

        'If e.Header.Column.Style = UltraWinGrid.ColumnStyle.CheckBox Then
        '    For Each aRow In e.Rows
        '        aRow.Cells(e.Header.Column.Index).Value = IIf((e.CheckState.Equals(CheckState.Checked)), "1", "0")
        '        aRow.Update()
        '    Next
        'End If
    End Sub

    ''' <summary>
    ''' Sub to tell the Creation filter what Groups in a grid are to have the Checkbox
    ''' </summary>
    ''' <param name="grd"></param>
    ''' <remarks></remarks>
    Public Overloads Sub DisplayHeaderCheckBox(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal GroupKeyList() As String)

        Dim columnList As New Dictionary(Of String, String)

        If GroupKeyList IsNot Nothing Then
            For Each key As String In GroupKeyList
                If key.Length > 0 Then
                    columnList.Add(key, "0")
                End If
            Next
        End If

        ' Set the columns to receive the checkbox
        aCheckBoxOnGroupHeader_CreationFilter.GroupKeys = columnList
        grd.CreationFilter = aCheckBoxOnGroupHeader_CreationFilter
    End Sub



End Class

''' <summary>
''' This CreationFilter class will create a CheckBoxUIElement in each Group Header
''' in the grid defined in the GroupKeys collection. It will fire the 
''' HeaderCheckBoxClicked event whenever the CheckBox is clicked.
''' </summary>
''' <remarks></remarks>
Public Class CheckBoxOnGroupHeader_CreationFilter
    Implements IUIElementCreationFilter
    Public GroupKeys As New Dictionary(Of String, String)

    ' This event will fire when the CheckBox is clicked. 
    Public Event HeaderCheckBoxClicked(ByVal sender As Object, ByVal e As HeaderCheckBoxEventArgs)

    Public Sub AfterCreateChildElements(ByVal parent As Infragistics.Win.UIElement) Implements Infragistics.Win.IUIElementCreationFilter.AfterCreateChildElements
        ' Check for the HeaderUIElement
        If TypeOf parent Is UltraWinGrid.HeaderUIElement AndAlso TypeOf parent.SelectableItem Is UltraWinGrid.GroupHeader Then
            ' Get the actual GroupHeader that the HeaderUIElement is attached to
            Dim aGrpHeader As Infragistics.Win.UltraWinGrid.GroupHeader
            aGrpHeader = CType(parent, UltraWinGrid.HeaderUIElement).Header.Group.Header

            ' Only put the Checkbox in the Header of Groups in the GroupKeys dictionary
            If GroupKeys.ContainsKey(aGrpHeader.Group.Key) Then
                Dim aTextUIElement As TextUIElement
                Dim aCheckBoxUIElement As CheckBoxUIElement

                ' Since the grid sometimes re-uses UIElements, we need to check to make sure 
                ' the header does not already have a CheckBoxUIElement attached to it.
                ' If it does, we just get a reference to the existing CheckBoxUIElement,
                ' and reset its properties.
                aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                If aCheckBoxUIElement Is Nothing Then
                    ' Create a New CheckBoxUIElement
                    aCheckBoxUIElement = New CheckBoxUIElement(parent)
                End If

                ' Get the TextUIElement - this is where the text for the 
                ' Header is displayed. We need this so we can push it to the right
                ' in order to make room for the CheckBox
                aTextUIElement = CType(parent.GetDescendant(GetType(TextUIElement)), TextUIElement)

                ' Sanity check
                If aTextUIElement Is Nothing Then Exit Sub

                ' Get the Header and see if the Tag has been set. I the Tag is 
                ' set, we will assume it's the stored CheckState. This has to be
                ' done in order to maintain the CheckState when the grid repaints and
                ' UIElement are destroyed and recreated. 

                If GroupKeys(aGrpHeader.Group.Key) = "" Then
                    ' If the tag was nothing, this is probably the first time this 
                    ' HeaderRow is being displayed
                    GroupKeys(aGrpHeader.Group.Key) = "0"
                Else
                    aCheckBoxUIElement.CheckState = IIf(GroupKeys(aGrpHeader.Group.Key) = "1", CheckState.Checked, CheckState.Unchecked)
                End If

                If aCheckBoxUIElement.CheckState = CheckState.Checked Then
                    aGrpHeader.Group.Header.Appearance.ForeColor = Color.Red
                Else
                    aGrpHeader.Group.Header.Appearance.ForeColor = Color.Empty
                End If

                ' Hook the ElementClick of the CheckBoxUIElement
                AddHandler aCheckBoxUIElement.ElementClick, AddressOf aCheckBoxUIElement_ElementClick

                ' Add the CheckBoxUIElement to the HeaderUIElement
                parent.ChildElements.Add(aCheckBoxUIElement)

                ' Position the CheckBoxUIElement. The number 3 here is used for 3
                ' pixels of padding between the CheckBox and the side of the header
                ' The CheckBox is shifted down slightly so it is centered in the header
                aCheckBoxUIElement.Rect = New Rectangle(parent.Rect.X + 3, parent.Rect.Y + ((parent.Rect.Height - aCheckBoxUIElement.CheckSize.Height) / 2), aCheckBoxUIElement.CheckSize.Width, aCheckBoxUIElement.CheckSize.Height)

                ' Push the TextUIElement to the right a little to make 
                ' room for the CheckBox. 3 pixels of padding are used again. 
                aTextUIElement.Rect = New Rectangle(aCheckBoxUIElement.Rect.Right + 3, aTextUIElement.Rect.Y, parent.Rect.Width - (aCheckBoxUIElement.Rect.Right - parent.Rect.X), aTextUIElement.Rect.Height)
            Else
                ' If the column is not a boolean column, we do not want to have a checkbox in it
                ' Since UIElements can be reused by the grid, there is a chance that one of the
                ' HeaderUIElements that we added a checkbox to for a boolean column header
                ' will be reused in a column that is not boolean.  In this case, we must remove
                ' the checkbox so that it will not appear in an inappropriate column header.
                Dim aCheckBoxUIElement As CheckBoxUIElement
                aCheckBoxUIElement = parent.GetDescendant(GetType(CheckBoxUIElement))

                If Not aCheckBoxUIElement Is Nothing Then
                    parent.ChildElements.Remove(aCheckBoxUIElement)
                    aCheckBoxUIElement.Dispose()
                End If
            End If
        End If
    End Sub

    Public Function BeforeCreateChildElements(ByVal parent As Infragistics.Win.UIElement) As Boolean Implements Infragistics.Win.IUIElementCreationFilter.BeforeCreateChildElements
        ' Don't need to do anything here.
        Return False
    End Function

    Private Sub aCheckBoxUIElement_ElementClick(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)
        ' Get the CheckBoxUIElement that was clicked
        Dim aCheckBoxUIElement As CheckBoxUIElement = CType(e.Element, CheckBoxUIElement)

        ' Get the Header associated with this particular element
        Dim aHeaderUIElement As UltraWinGrid.HeaderUIElement = CType(aCheckBoxUIElement.GetAncestor(GetType(UltraWinGrid.HeaderUIElement)), UltraWinGrid.HeaderUIElement)
        Dim aGrpHeader As Infragistics.Win.UltraWinGrid.GroupHeader = CType(aHeaderUIElement.GetContext(GetType(Infragistics.Win.UltraWinGrid.GroupHeader)), Infragistics.Win.UltraWinGrid.GroupHeader)


        ' Set the Tag on the Header to the new CheckState
        GroupKeys(aGrpHeader.Group.Key) = IIf(aCheckBoxUIElement.CheckState = CheckState.Checked, "1", "0")

        ' So that we can apply various changes only to the relevant Rows collection that the header belongs to
        Dim hRows As UltraWinGrid.RowsCollection = CType(aHeaderUIElement.GetContext(GetType(UltraWinGrid.RowsCollection)), UltraWinGrid.RowsCollection)

        ' Raise an event so the programmer can do something when the CheckState changes
        RaiseEvent HeaderCheckBoxClicked(Me, New HeaderCheckBoxEventArgs(aGrpHeader, aCheckBoxUIElement.CheckState, hRows))
    End Sub

    ' EventArgs used for the HeaderCheckBoxClicked event. This event has to pass in the CheckState and the Header
    Public Class HeaderCheckBoxEventArgs
        Inherits EventArgs

        Public Sub New(ByVal Header As Infragistics.Win.UltraWinGrid.GroupHeader, ByVal CheckState As CheckState, ByRef Rows As UltraWinGrid.RowsCollection)
            mvarHeader = Header
            mvarCheckState = CheckState
            mvarRowsCollection = Rows
        End Sub

        Private mvarRowsCollection As UltraWinGrid.RowsCollection
        Private mvarHeader As Infragistics.Win.UltraWinGrid.GroupHeader
        Private mvarCheckState As CheckState

        ' Expose the rows collection for the specific row island that the header belongs to
        Public ReadOnly Property Rows() As UltraWinGrid.RowsCollection
            Get
                Return mvarRowsCollection
            End Get
        End Property

        Public ReadOnly Property Header() As Infragistics.Win.UltraWinGrid.GroupHeader
            Get
                Return mvarHeader
            End Get
        End Property

        Public Property CheckState() As CheckState
            Get
                Return mvarCheckState
            End Get
            Set(ByVal Value As CheckState)
                mvarCheckState = Value
            End Set
        End Property
    End Class

End Class