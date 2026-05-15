Public Class RSFBUDR2
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim BRAND_CODE As String
    Dim OPS_YYYY As String
    Dim OPS_YYYY_LY As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)

            ASCMAIN1.sql = "Select *" & vbCrLf _
                & " from ICTCOLL1 where BRAND_CODE = :PARM1 or :PARM1 is Null" & vbCrLf _
                & " order by COLLECTION_CODE"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, SELL_CODE, ARTCUST2.CUST_STORE_NAME, " _
            & " NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _
            & ", CUST_STORE_CITY, CUST_STORE_STATE" _
            & " from ARTCUST2 where CUST_CODE = :PARM1 or :PARM1 is Null"
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select RSTBUDR1.* from RSTBUDR1" _
            & " where OPS_YYYY = :PARM1 and CUST_CODE = :PARM2" _
            & " and COLLECTION_CODE in " _
            & " (Select COLLECTION_CODE FROM ICTCOLL1 where BRAND_CODE = :PARM3)"
            Create_TDA(.Tables.Add, "RSTBUDR1", "**", 0, True, "VVV", 5)
            .Tables("RSTBUDR1").Columns("ITEM_CATGY_CODE").DefaultValue = "E"
            For M As Integer = 1 To 12
                With .Tables("RSTBUDR1").Columns("BUDGET_P" & Format(M, "00"))
                    .DataType = GetType(System.Int32)
                End With
            Next
            .Relations.Add("ARTCUSTX_RSTBUDR1", _
                    New DataColumn() {.Tables("ARTCUSTX").Columns("CUST_CODE"), .Tables("ARTCUSTX").Columns("CUST_STORE_NO")}, _
                    New DataColumn() {.Tables("RSTBUDR1").Columns("CUST_CODE"), .Tables("RSTBUDR1").Columns("CUST_STORE_NO")})
            With .Tables("RSTBUDR1").Columns
                .Add("TOTAL", GetType(System.Int32), _
                      "ISNULL(BUDGET_P01,0)+ISNULL(BUDGET_P02,0)+ISNULL(BUDGET_P03,0)+" _
                    & "ISNULL(BUDGET_P04,0)+ISNULL(BUDGET_P05,0)+ISNULL(BUDGET_P06,0)+" _
                    & "ISNULL(BUDGET_P07,0)+ISNULL(BUDGET_P08,0)+ISNULL(BUDGET_P09,0)+" _
                    & "ISNULL(BUDGET_P10,0)+ISNULL(BUDGET_P11,0)+ISNULL(BUDGET_P12,0)")
                .Add("CUST_STORE_LOCATION", GetType(System.String), "PARENT.CUST_STORE_LOCATION")
                .Add("SELL_CODE", GetType(System.String), "PARENT.SELL_CODE")
                .Add("CUST_STORE_CITY", GetType(System.String), "PARENT.CUST_STORE_CITY")
                .Add("CUST_STORE_STATE", GetType(System.String), "PARENT.CUST_STORE_STATE")
            End With

        End With

        grdRSTBUDR1.DataSource = dst.Tables("RSTBUDR1")

        Create_Summary(grdRSTBUDR1, "CUST_STORE_NO", "Count")
        For M As Integer = 1 To 12
            Create_Summary(grdRSTBUDR1, "BUDGET_P" & Format(M, "00"), , , "###,##0")
        Next
        Create_Summary(grdRSTBUDR1, "TOTAL", , , "###,##0")

        With grdRSTBUDR1.DisplayLayout.Bands("RSTBUDR1")
            For Each COLUMN_NAME As String In New String() _
                {"OPS_YYYY", "COLLECTION_CODE", "ITEM_CATGY_CODE", "CUST_CODE", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                If New String() {"OPS_YYYY", "CUST_CODE", "CUST_STORE_NO", "COLLECTION_CODE", "ITEM_CATGY_CODE", "TOTAL", "CUST_STORE_LOCATION", "SELL_CODE", "CUST_STORE_CITY", "CUST_STORE_STATE"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = System.Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs ' New String() {"2008", "2009", "2010"}

        optITEM_CATGY_CODE.ValueList = ASCMAIN1.ValueListFor("ITEM_CATGY_CODE", , New String() {":", "*:All"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")
                Validate_Code("BRAND_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & Absx1.cbeFor("OPS_YYYY").Value) Then
                        Exit Sub
                    End If
                    If optCalendar.Value = "O" Then
                        If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & Format(Val(Absx1.cbeFor("OPS_YYYY").Value & "") - 1, "0000")) Then
                            Exit Sub
                        End If
                    End If
                    If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"CUST_CODE", "CUST_STORE_NO"}).Select("")
                    Dim CUST_CODE As String = row.Item(0)
                    Dim CUST_STORE_NO As String = row.Item(1)
                    If dst.Tables("ARTCUSTX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Store (" & CUST_CODE & "," & CUST_STORE_NO & ")"
                    End If
                Next
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"COLLECTION_CODE"}).Select("")
                    Dim COLLECTION_CODE As String = row.Item(0)
                    If dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                    End If
                Next
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"OPS_YYYY"}).Select("")
                    Dim OPS_YYYY As String = row.Item(0)
                    If OPS_YYYY <> Absx1.cbeFor("OPS_YYYY").Value & "" Then
                        EMsg &= vbCr & "Invalid Budget Year (" & OPS_YYYY & ")"
                    End If
                Next

            Case "Load Stores"
                If optITEM_CATGY_CODE.Value = "A" Or optCOLLECTION_CODE.Value = "A" Then
                    EMsg &= vbCr & "Cannot Load Stores unless Individual Values are chosen for Collection and Category"
                End If

            Case "Import from XLS"

                OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Value & ""
                OPS_YYYY_LY = Format(Val(OPS_YYYY) - 1, "0000")

                If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & OPS_YYYY) Then
                    Exit Sub
                End If
                If optCalendar.Value = "O" Then
                    If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & OPS_YYYY_LY) Then
                        Exit Sub
                    End If
                End If

                Absx1.txtFor("CUST_CODE").Text = ""
                Absx1.txtFor("BRAND_CODE").Text = ""

                dst.Tables("RSTBUDR1").Rows.Clear()
                Fill_Records("ARTCUSTX", "")
                Fill_Records("ICTCOLL1", "")

                grdRSTBUDR1.Text = "Retail Sales Budgets, by Store/Month, for " & Absx1.cbeFor("OPS_YYYY").Value
                grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdRSTBUDR1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdRSTBUDR1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes

                With grdRSTBUDR1.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"OPS_YYYY", "CUST_CODE", "CUST_STORE_NO", "ITEM_CATGY_CODE", "COLLECTION_CODE"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                End With

                Set_Month_Headings(OPS_YYYY)
                grdRSTBUDR1.Visible = True
                grdRSTBUDR1.DisplayLayout.Bands(0).Columns("ITEM_CATGY_CODE").Hidden = True

                Excel_Import_SG(grdRSTBUDR1)

                grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdRSTBUDR1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTBUDR1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

                With grdRSTBUDR1.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"OPS_YYYY", "CUST_CODE", "CUST_STORE_NO", "ITEM_CATGY_CODE", "COLLECTION_CODE"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    Next
                End With

                If dst.Tables("RSTBUDR1").Rows.Count = 0 Then
                    ASCMAIN1.MultiTask_Release()
                    EMsg &= vbCr & "No Budget Records Imported"
                Else
                    Sort_grdColumns(grdRSTBUDR1, "OPS_YYYY,CUST_CODE,COLLECTION_CODE,CUST_STORE_NO")
                End If


            Case "Show All"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Data")

                Absx1.txtFor("CUST_CODE").Text = ""
                Absx1.txtFor("BRAND_CODE").Text = ""

                OPS_YYYY = Absx1.cbeFor("OPS_YYYY").Value & ""

                dst.Tables("RSTBUDR1").Rows.Clear()
                Fill_Records("ARTCUSTX", "")
                Fill_Records("ICTCOLL1", "")

                grdRSTBUDR1.Text = "Retail Sales Budgets, by Store/Month, for " & OPS_YYYY
                Set_Month_Headings(OPS_YYYY)
                grdRSTBUDR1.Visible = True

                grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdRSTBUDR1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdRSTBUDR1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

                With grdRSTBUDR1.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"OPS_YYYY", "CUST_CODE", "CUST_STORE_NO", "ITEM_CATGY_CODE", "COLLECTION_CODE"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                    Next
                End With

                If optCalendar.Value = "R" Then
                    ASCMAIN1.sql = "Select * from RSTBUDR1 where OPS_YYYY = '" & OPS_YYYY & "'"
                Else
                    ASCMAIN1.sql = Get_SQL_Operational_Calendar(OPS_YYYY)
                End If

                Fill_Records("RSTBUDR1", "", True, ASCMAIN1.sql)
                Sort_grdColumns(grdRSTBUDR1, "OPS_YYYY,CUST_CODE,COLLECTION_CODE,CUST_STORE_NO")

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
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

            Case "Load Stores"

                Dim ITEM_CATGY_CODE As String = optITEM_CATGY_CODE.Value
                Dim COLLECTION_CODE As String = cbeCOLLECTION_CODE.Value & ""
                If COLLECTION_CODE = "" Then
                    MsgBox("Cannot Load Stores unless a Single Collection is Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                For Each row As DataRow In dst.Tables("ARTCUSTX").Rows
                    Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                    Dim rowRSTBUDR1 As DataRow = dst.Tables("RSTBUDR1").Rows.Find(New Object() _
                    {OPS_YYYY, COLLECTION_CODE, ITEM_CATGY_CODE, CUST_CODE, CUST_STORE_NO})
                    If rowRSTBUDR1 Is Nothing Then
                        rowRSTBUDR1 = dst.Tables("RSTBUDR1").NewRow
                        rowRSTBUDR1.Item("OPS_YYYY") = OPS_YYYY
                        rowRSTBUDR1.Item("COLLECTION_CODE") = COLLECTION_CODE
                        rowRSTBUDR1.Item("ITEM_CATGY_CODE") = ITEM_CATGY_CODE
                        rowRSTBUDR1.Item("CUST_CODE") = CUST_CODE
                        rowRSTBUDR1.Item("CUST_STORE_NO") = CUST_STORE_NO
                        dst.Tables("RSTBUDR1").Rows.Add(rowRSTBUDR1)
                    End If
                Next

            Case "Import from XLS"
                EntryMode = "I"
                ' Load_Record()
                Mode_Settings(True)

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
                    .Items("Load Stores").Settings.Enabled = iScreenMode
                    .Items("Load Stores").Visible = (EntryMode = "L")
                    .Items("Import from XLS").Visible = Not tf

                    .Items("Show All").Visible = Not tf

                End With

                .Groups("Display Options").Visible = ScreenMode And (EntryMode = "L")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdRSTBUDR1.Visible = ScreenMode

        With grdRSTBUDR1.DisplayLayout.Bands(0)
            .Columns("OPS_YYYY").Hidden = ScreenMode And (EntryMode = "L")
            .Columns("CUST_CODE").Hidden = ScreenMode And (EntryMode = "L")
            .Columns("COLLECTION_CODE").Hidden = ScreenMode And (EntryMode = "L")
        End With

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUSTX", "RSTBUDR1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        BRAND_CODE = ""
        OPS_YYYY = ""

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
        Absx1.cbeFor("OPS_YYYY").Value = Now.Year
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        BRAND_CODE = HFs("BRAND_CODE")
        OPS_YYYY = HFs("OPS_YYYY")
        OPS_YYYY_LY = Format(Val(OPS_YYYY) - 1, "0000")

        EnforceConstraints(False)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        Fill_Records("ARTCUSTX", CUST_CODE)

        If optCalendar.Value = "R" Then
            Fill_Records("RSTBUDR1", New String() {OPS_YYYY, CUST_CODE, BRAND_CODE})
        Else
            ASCMAIN1.sql = Get_SQL_Operational_Calendar(OPS_YYYY, CUST_CODE, BRAND_CODE)
            Fill_Records("RSTBUDR1", "", True, ASCMAIN1.sql)
        End If

        For Each row As DataRow In ASCDATA1.SelectDistinct("RSTBUDR1", "CUST_STORE_NO").Rows
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowARTCUSTX Is Nothing Then
                ' WHY ARE WE ADDING STORES THAT DO NOT EXIST?
                rowARTCUSTX = dst.Tables("ARTCUSTX").NewRow
                rowARTCUSTX.Item("CUST_CODE") = CUST_CODE
                rowARTCUSTX.Item("CUST_STORE_NO") = CUST_STORE_NO
                dst.Tables("ARTCUSTX").Rows.Add(rowARTCUSTX)
            End If
        Next

        Fill_Records("ICTCOLL1", BRAND_CODE)

        EnforceConstraints(True)

        Sort_grdColumns(grdRSTBUDR1, "CUST_STORE_NO")
        optITEM_CATGY_CODE.Value = "E"

        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "' order by COLLECTION_CODE"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        'cbeCOLLECTION_CODE.DataSource = dst.Tables("ICTCOLL1")
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Set_Month_Headings(OPS_YYYY)

        Setup_grdRSTBUDR1()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim sqlCB As String = "" _
            & " and CUST_CODE = '" & CUST_CODE & "'" _
            & " and COLLECTION_CODE IN " _
            & " (Select COLLECTION_CODE FROM ICTCOLL1 " _
            & " where BRAND_CODE = '" & BRAND_CODE & "')"

        Dim YYYYs As String = "'" & OPS_YYYY & "'"
        If optCalendar.Value = "O" Then
            YYYYs &= ",'" & OPS_YYYY_LY & "'"
        End If

        Dim sql_Delete As String = "Delete from RSTBUDR1" _
            & " where OPS_YYYY in (" & YYYYs & ")" _
            & IIf(EntryMode = "I", "", sqlCB)

        If optCalendar.Value = "O" Then
            Dim J As Integer = dst.Tables("RSTBUDR1").Columns("BUDGET_P01").Ordinal

            For Each row As DataRow In dst.Tables("RSTBUDR1").Select("OPS_YYYY = '" & OPS_YYYY & "'")
                Dim rowLY As DataRow = dst.Tables("RSTBUDR1").NewRow
                For Each C As String In New String() {"COLLECTION_CODE", "ITEM_CATGY_CODE", "CUST_CODE", "CUST_STORE_NO"}
                    rowLY.Item(C) = row.Item(C)
                Next
                rowLY.Item("OPS_YYYY") = OPS_YYYY_LY
                rowLY.Item("BUDGET_P12") = row.Item("BUDGET_P01")
                dst.Tables("RSTBUDR1").Rows.Add(rowLY)

                For I As Integer = 1 To 11
                    row.Item(J + I - 1) = row(J + I)
                Next
            Next

            ASCMAIN1.sql = "Select * from RSTBUDR1" _
                & " where OPS_YYYY in (" & YYYYs & ")" _
                & IIf(EntryMode = "I", "", sqlCB)
            Dim tbl As DataTable = ASCDATA1.GetDataTable

            For Each row As DataRow In tbl.Select("")
                Dim write_record As Boolean = False
                Dim B(12) As Decimal
                Dim B1 As Integer = 1
                Dim B2 As Integer = 11

                If row.Item("OPS_YYYY") = OPS_YYYY_LY Then
                    B1 = 12
                    B2 = 12
                End If

                For I As Integer = B1 To B2
                    Dim BUDGET As Decimal = Val(row.Item("BUDGET_P" & Format(I, "00")) & "")
                    If BUDGET <> 0 Then
                        B(I) = BUDGET
                        write_record = True
                    End If
                Next

                If write_record Then
                    Dim rowRSTBUDR1 As DataRow = dst.Tables("RSTBUDR1") _
                    .Rows.Find(New String() { _
                      row.Item("OPS_YYYY"), _
                      row.Item("COLLECTION_CODE"), _
                      row.Item("ITEM_CATGY_CODE"), _
                      row.Item("CUST_CODE"), _
                      row.Item("CUST_STORE_NO")})
                    If rowRSTBUDR1 Is Nothing Then
                        rowRSTBUDR1 = dst.Tables("RSTBUDR1").Rows.Add(New String() { _
                          row.Item("OPS_YYYY"), _
                          row.Item("COLLECTION_CODE"), _
                          row.Item("ITEM_CATGY_CODE"), _
                          row.Item("CUST_CODE"), _
                          row.Item("CUST_STORE_NO")})
                    End If
                    For I As Integer = B1 To B2
                        If B(I) <> 0 Then
                            row.Item("BUDGET_P" & Format(I, "00")) = B(I)
                        End If
                    Next
                End If
            Next

            For M As Integer = 1 To 12
                Dim D As Date = CDate(Format(M, "00") & "/01/" & OPS_YYYY).AddMonths(1)
                Dim LEGEND As String = Format(D, "MMM")
                With grdRSTBUDR1.DisplayLayout.Bands(0).Columns("BUDGET_P" & Format(M, "00"))
                    .Header.Caption = LEGEND
                End With
            Next

        End If

        Update_Record_TDA("RSTBUDR1", sql_Delete)

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
        Load_Popup_Menu(grdRSTBUDR1, "BBBB", "Clear Column", "Copy Value", "Export Spring", "Export Fall")
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
                Case "grdRSTBUDR1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Export Spring"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (Not ScreenMode)
                    tlb_btn = DirectCast(tlb_pop.Tools("Export Fall"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (Not ScreenMode)

                    tlb_btn = DirectCast(tlb_pop.Tools("Clear Column"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Value"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")

                    If grdRSTBUDR1.Tag = "" Then
                        'e.Cancel = True
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Export Fall"
                Create_Pivot("F")
            Case "Export Spring"
                Create_Pivot("S")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Clear Column"
                Dim COLUMN_NAME As String = grdRSTBUDR1.Tag
                If COLUMN_NAME = "" Then Exit Sub
                If COLUMN_NAME = "CUST_STORE_NO" Then Exit Sub
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = DBNull.Value
                Next
            Case "Copy Value"
                Dim COLUMN_NAME As String = grdRSTBUDR1.Tag
                If COLUMN_NAME = "" Then Exit Sub
                If grdRSTBUDR1.ActiveRow Is Nothing OrElse grdRSTBUDR1.ActiveRow.IsAddRow OrElse Not grdRSTBUDR1.ActiveRow.IsDataRow Then Exit Sub
                Dim COPY_VALUE As String = grdRSTBUDR1.ActiveRow.Cells(COLUMN_NAME).Value
                For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
                    row.Item(COLUMN_NAME) = COPY_VALUE
                Next

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
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
            Case "BRAND_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
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

            Case "BRAND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "grdRSTBUDR1"

    Private Sub grdRSTBUDR1_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTBUDR1.AfterExitEditMode
        With grdRSTBUDR1
            Select Case .ActiveCell.Column.Key
                Case "CUST_STORE_NO"
                    Dim CUST_STORE_NO As String = .ActiveCell.Text
                    If CUST_STORE_NO <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(CUST_STORE_NO, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTBUDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTBUDR1.AfterRowActivate
        With grdRSTBUDR1.DisplayLayout.Bands(0)
            If grdRSTBUDR1.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ITEM_CATGY_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ITEM_CATGY_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdRSTBUDR1_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdRSTBUDR1.BeforeExitEditMode
        'e.Cancel = True
    End Sub

    Private Sub grdRSTBUDR1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTBUDR1.BeforeRowUpdate
        With grdRSTBUDR1
            'If Val(e.Row.Cells("ORDR_QTY").Text) < 0 Then
            '    MsgBox("Invalid Value entered for Order Qty (" & e.Row.Cells("ORDR_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If

            'If e.Cancel Then
            '    e.Row.CancelUpdate()
            'End If

            If Not e.Cancel Then
                If e.Row.Cells("OPS_YYYY").Text = "" And ScreenMode Then
                    .ActiveRow.Cells("OPS_YYYY").Value = Absx1.CtlFor("OPS_YYYY").Text
                    .ActiveRow.Cells("CUST_CODE").Value = Absx1.CtlFor("CUST_CODE").Text
                    .ActiveRow.Cells("ITEM_CATGY_CODE").Value = optITEM_CATGY_CODE.Value
                    .ActiveRow.Cells("COLLECTION_CODE").Value = cbeCOLLECTION_CODE.Value
                End If
            End If
        End With
    End Sub

    Private Sub grdRSTBUDR1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTBUDR1.ClickCellButton
        Select Case grdRSTBUDR1.ActiveCell.Column.Key
            Case "COLLECTION_CODE"
                grdClickCellButton(grdRSTBUDR1)
            Case "CUST_CODE"
                grdClickCellButton(grdRSTBUDR1)
            Case "CUST_STORE_NO"
                'grdClickCellButton(grdRSTBUDR1, "CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Text & "'")
                grdClickCellButton(grdRSTBUDR1, "CUST_CODE = '" & HFs("CUST_CODE") & "'")
        End Select
    End Sub

    Private Sub grdRSTBUDR1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTBUDR1.InitializeLayout
        'e.Layout.Override.AllowMultiCellOperations = Infragistics.Win.UltraWinGrid.AllowMultiCellOperation.All
        ' e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.CellSelect
    End Sub

#End Region

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdRSTBUDR1()
    End Sub

    Private Sub optBP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optITEM_CATGY_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdRSTBUDR1()
    End Sub

    Sub Setup_grdRSTBUDR1()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdRSTBUDR1.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdRSTBUDR1.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").Hidden = (optCOLLECTION_CODE.Value <> "A")
            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdRSTBUDR1.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("CUST_STORE_NO", False)

            .Columns("ITEM_CATGY_CODE").Hidden = (optITEM_CATGY_CODE.Value <> "A")
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdRSTBUDR1.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            allow_modifications = False
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        If optITEM_CATGY_CODE.Value = "*" Then
            allow_modifications = False
        Else
            sql &= " and ITEM_CATGY_CODE = '" & optITEM_CATGY_CODE.Value & "'"
        End If

        DVW.RowFilter = Mid(sql, 5)
        grdRSTBUDR1.Text = "Retail Sales Budgets, by Store/Month, for " & Absx1.cbeFor("OPS_YYYY").Value & " - " & COLLS
        If allow_modifications Then
            grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdRSTBUDR1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdRSTBUDR1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Else
            grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdRSTBUDR1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdRSTBUDR1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        End If
        'grdRSTBUDR1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        grdRSTBUDR1.DisplayLayout.Bands(0).Columns("ITEM_CATGY_CODE").Hidden = (optITEM_CATGY_CODE.Value <> "*")
        ASCMAIN1.Add_Value_List(grdRSTBUDR1, "ITEM_CATGY_CODE")

    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdRSTBUDR1()
    End Sub

    Sub Set_Month_Headings(OPS_YYYY As String)
        Dim MM As Integer = 1
        If optCalendar.Value = "O" Then
            MM = 0
        End If

        For M As Integer = 1 To 12
            Dim D As Date = CDate(Format(M, "00") & "/01/" & OPS_YYYY).AddMonths(MM)
            Dim LEGEND As String = Format(D, "MMM") & "'" & Format(D, "yy")
            With grdRSTBUDR1.DisplayLayout.Bands(0).Columns("BUDGET_P" & Format(M, "00"))
                .Header.Caption = LEGEND
                .Width = 60
            End With
        Next
    End Sub

    Sub Create_Pivot(SEASON_TYPE As String)

        Dim SQLW As String = ""

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & "_" & SEASON_TYPE & ".xlsX"

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim SheetName As String = "MainData"
        ws = wb.Worksheets(SheetName)

        Dim DataTable As DataTable = dst.Tables("RSTBUDR1")

        Dim Formulae As New List(Of Integer)
        Dim CUST_CODE As String = ""
        Dim CUST_STORE_NO As String = ""
        Dim COLLECTION_CODE As String = ""

        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowICTCOLL1 As DataRow = Nothing

        ASCMAIN1.sql = "Select * from ARTCUST1"
        Fill_Records("ARTCUST1", "", True, ASCMAIN1.sql)


        ASCMAIN1.Progress("-", "History")

        Dim FYP As String = Absx1.cbeFor("OPS_YYYY").Value & IIf(SEASON_TYPE = "S", "02", "08")

        Dim sqlP As String = ""
        For i As Integer = 1 To 6
            Dim j As Integer = i + 1 + IIf(SEASON_TYPE = "F", 6, 0)
            If j = 13 Then j = 1
            sqlP &= ", Sum (Decode(SUBSTR(GLTPARM2.OPS_YYYYMM,5,2),'" & Format(j, "00") & "',RETAIL_SALES,0)) RETAIL_SALES_" & Format(i, "00") & vbCrLf
        Next
        Dim sqlH As String = "" _
            & "Select RSTRETL2.CUST_CODE, RSTRETL2.CUST_STORE_NO, RSTRETL2.COLLECTION_CODE" & vbCrLf _
            & ", Sum (RSTRETL2.RETAIL_SALES) RETAIL_SALES" & vbCrLf _
            & sqlP _
            & " from RSTRETL2,GLTPARM2" & vbCrLf _
            & " where RSTRETL2.OPS_YYYYPP between :PARM1 and :PARM2" & vbCrLf _
            & "   and RSTRETL2.RETAIL_SALES <> 0" & vbCrLf _
            & "   and GLTPARM2.OPS_YYYYPP = RSTRETL2.OPS_YYYYPP" & vbCrLf _
            & " group by RSTRETL2.CUST_CODE, RSTRETL2.CUST_STORE_NO, RSTRETL2.COLLECTION_CODE"

        Dim tblLY As DataTable = ASCDATA1.GetDataTable(sqlH, "LY", 3, True, 0, "VV", _
                                                       New String() {ASCMAIN1.Period_Calc(FYP, -12), _
                                                                     ASCMAIN1.Period_Calc(FYP, -12 + 6)})
        'tblLY.PrimaryKey = New DataColumn() {tblLY.Columns("CUST_CODE"), tblLY.Columns("CUST_STORE_NO"), tblLY.Columns("COLLECTION_CODE")}

        Dim tbl2LY As DataTable = ASCDATA1.GetDataTable(sqlH, "2LY", 3, True, 0, "VV", _
                                                       New String() {ASCMAIN1.Period_Calc(FYP, -12 - 12), _
                                                                     ASCMAIN1.Period_Calc(FYP, -12 - 12 + 6)})
        'tbl2LY.PrimaryKey = New DataColumn() {tbl2LY.Columns("CUST_CODE"), tbl2LY.Columns("CUST_STORE_NO"), tbl2LY.Columns("COLLECTION_CODE")}

        Dim iRx As Integer = 8
        Dim r As Integer = 0 ' since we are using XLS Automation
        Dim c As Integer

        For Each row As DataRow In DataTable.Select("", "CUST_CODE,COLLECTION_CODE")
            r += 1
            ASCMAIN1.Progress("-", r)
            c = 0
            'Formulae.Clear()
            ' ws.Range("A" & CStr(3 + r) & ":AL" & CStr(3 + r)).Value2 = row.ItemArray
            If CUST_CODE <> row.Item("CUST_CODE") Then
                CUST_CODE = row.Item("CUST_CODE")
                rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            End If
            If COLLECTION_CODE <> row.Item("COLLECTION_CODE") Then
                COLLECTION_CODE = row.Item("COLLECTION_CODE")
                rowICTCOLL1 = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
            End If
            CUST_STORE_NO = row.Item("CUST_STORE_NO")
            Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})

            Dim BUDGET As Decimal = 0
            For I As Integer = 1 To 6
                Dim J As Integer = I + IIf(SEASON_TYPE = "F", 6, 0)
                BUDGET += Val(row.Item("BUDGET_P" & Format(J, "00")) & "")
            Next

            Dim rowLY As DataRow = tblLY.Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, COLLECTION_CODE})
            Dim RETAIL_SALES_LY() As Decimal = New Decimal() {0, 0, 0, 0, 0, 0, 0}
            If rowLY IsNot Nothing Then
                For i As Integer = 1 To 6
                    RETAIL_SALES_LY(i) = Val(rowLY.Item("RETAIL_SALES_" & Format(i, "00")) & "")
                Next
                'RETAIL_SALES_LY = Val(rowLY.Item("RETAIL_SALES") & "")
            End If
            Dim row2LY As DataRow = tbl2LY.Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, COLLECTION_CODE})
            Dim RETAIL_SALES_2LY() As Decimal = New Decimal() {0, 0, 0, 0, 0, 0, 0}
            If row2LY IsNot Nothing Then
                For i As Integer = 1 To 6
                    RETAIL_SALES_2LY(i) = Val(row2LY.Item("RETAIL_SALES_" & Format(i, "00")) & "")
                Next
                'RETAIL_SALES_2LY = Val(row2LY.Item("RETAIL_SALES") & "")
            End If

            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = CUST_CODE & "-" & CUST_STORE_NO
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = CUST_CODE
            c += 1 : ws.Cells(iRx + r, c).Value2 = rowARTCUST1.Item("CUST_NAME")
            c += 1 : ws.Cells(iRx + r, c).Value2 = "Field"
            c += 1 : ws.Cells(iRx + r, c).Value2 = rowARTCUSTX.Item("CUST_STORE_NAME")
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = "IPLB"
            c += 1 : ws.Cells(iRx + r, c).Value2 = rowICTCOLL1.Item("COLLECTION_GENDER")
            c += 1 : ws.Cells(iRx + r, c).Value2 = COLLECTION_CODE
            c += 1 : ws.Cells(iRx + r, c).Value2 = rowICTCOLL1.Item("COLLECTION_NAME")
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = ""
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(1)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(2)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(3)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(4)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(5)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_2LY(6)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(1)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(2)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(3)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(4)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(5)
            c += 1 : ws.Cells(iRx + r, c).Value2 = RETAIL_SALES_LY(6)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = BUDGET
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : ws.Cells(iRx + r, c).Value2 = ""
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
            c += 1 : If r = 1 Then Formulae.Add(c)
        Next

        ASCMAIN1.Progress("-", "Formulae")
        For Each Cx As Integer In Formulae
            Dim XC As String = Excel_Cell(iRx + 1, Cx)
            Dim XD As String = Excel_Cell(iRx + DataTable.Rows.Count, Cx)
            xlSourceRange = ws.Range(XC, XC)
            xlDestRange = ws.Range(XC, XD)
            xlSourceRange.Copy(xlDestRange)
        Next

        ASCMAIN1.Progress("-", "Pivot")
        wb.Names.Add("PivotBase", "=" & SheetName & "!" & Excel_Cell(iRx, 1, 3) & CStr(iRx) & ":" & Excel_Cell(iRx + DataTable.Rows.Count, c + 1, 3))

        ' excel.Run("ResetData")

        '    Dim XLS_filename As String = Save_WorkBook_and_ReleaseCOM("PO_Log")
        'End Sub

        'Function Save_WorkBook_and_ReleaseCOM(filename_pfx As String) As String

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = SEASON_TYPE & Absx1.cbeFor("OPS_YYYY").Value & "_Planner"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsX"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

        'Return ASCMAIN1.Folders("Work") & XLS_FILENAME

    End Sub

    Private Sub optCalendar_ValueChanged(sender As Object, e As EventArgs) Handles optCalendar.ValueChanged

    End Sub

    Function Get_SQL_Operational_Calendar(YYYY As String, _
                                          Optional CUST_CODE As String = "", _
                                          Optional BRAND_CODE As String = "") As String

        Dim SQLC As String = ""
        If CUST_CODE <> "" Then
            SQLC &= " and CUST_CODE = '" & CUST_CODE & "'"
        End If
        If BRAND_CODE <> "" Then
            SQLC &= " and COLLECTION_CODE in (Select COLLECTION_CODE from ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "')"
        End If

        Dim Sql As String = "" _
            & "Select OPS_YYYY,COLLECTION_CODE,ITEM_CATGY_CODE,CUST_CODE,CUST_STORE_NO" & vbCrLf _
            & ", SUM (BUDGET_P01) BUDGET_P01" & vbCrLf _
            & ", SUM (BUDGET_P02) BUDGET_P02" & vbCrLf _
            & ", SUM (BUDGET_P03) BUDGET_P03" & vbCrLf _
            & ", SUM (BUDGET_P04) BUDGET_P04" & vbCrLf _
            & ", SUM (BUDGET_P05) BUDGET_P05" & vbCrLf _
            & ", SUM (BUDGET_P06) BUDGET_P06" & vbCrLf _
            & ", SUM (BUDGET_P07) BUDGET_P07" & vbCrLf _
            & ", SUM (BUDGET_P08) BUDGET_P08" & vbCrLf _
            & ", SUM (BUDGET_P09) BUDGET_P09" & vbCrLf _
            & ", SUM (BUDGET_P10) BUDGET_P10" & vbCrLf _
            & ", SUM (BUDGET_P11) BUDGET_P11" & vbCrLf _
            & ", SUM (BUDGET_P11) BUDGET_P12" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select RSTBUDR1.OPS_YYYY,RSTBUDR1.COLLECTION_CODE,RSTBUDR1.ITEM_CATGY_CODE,RSTBUDR1.CUST_CODE,RSTBUDR1.CUST_STORE_NO" & vbCrLf _
            & ", 0 BUDGET_P01" _
            & ", RSTBUDR1.BUDGET_P01 BUDGET_P02" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P02 BUDGET_P03" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P03 BUDGET_P04" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P04 BUDGET_P05" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P05 BUDGET_P06" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P06 BUDGET_P07" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P07 BUDGET_P08" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P08 BUDGET_P09" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P09 BUDGET_P10" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P10 BUDGET_P11" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P11 BUDGET_P12" & vbCrLf _
            & "from RSTBUDR1 where OPS_YYYY = '" & YYYY & "'" & vbCrLf _
            & SQLC _
            & "union" & vbCrLf _
            & "Select RSTBUDR1.OPS_YYYY,RSTBUDR1.COLLECTION_CODE,RSTBUDR1.ITEM_CATGY_CODE,RSTBUDR1.CUST_CODE,RSTBUDR1.CUST_STORE_NO" & vbCrLf _
            & ", RSTBUDR1.BUDGET_P12 BUDGET_P01" & vbCrLf _
            & ", 0 BUDGET_P02" & vbCrLf _
            & ", 0 BUDGET_P03" & vbCrLf _
            & ", 0 BUDGET_P04" & vbCrLf _
            & ", 0 BUDGET_P05" & vbCrLf _
            & ", 0 BUDGET_P06" & vbCrLf _
            & ", 0 BUDGET_P07" & vbCrLf _
            & ", 0 BUDGET_P08" & vbCrLf _
            & ", 0 BUDGET_P09" & vbCrLf _
            & ", 0 BUDGET_P10" & vbCrLf _
            & ", 0 BUDGET_P11" & vbCrLf _
            & ", 0 BUDGET_P12" & vbCrLf _
            & "from RSTBUDR1 where OPS_YYYY = '" & Format(Val(YYYY) - 1, "0000") & "'" & vbCrLf _
            & SQLC _
            & ") group by OPS_YYYY,COLLECTION_CODE,ITEM_CATGY_CODE,CUST_CODE,CUST_STORE_NO"
        Return Sql

    End Function

    Overrides Function Excel_Import_Pre_Process_SG _
    (ByVal grd As UltraWinGrid.UltraGrid, dt As DataTable,
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing) As Int64

        Dim dtbad As DataTable = dt.Clone
        dtbad.Columns.Add("ERROR")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Budgets from XLS")

        Dim RowsMax As Int64 = dt.Rows.Count
        Dim r As Int64 = 0

        For Each row As DataRow In dt.Select("")
            r += 1
            If r Mod 100 = 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Budgets from XLS")
                RowsMax = dt.Rows.Count
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RowsMax))
            End If

            Try
                Dim rowRSTBUDR1 As DataRow = dst.Tables("RSTBUDR1").NewRow
                With rowRSTBUDR1
                    For Each C As String In New String() {"OPS_YYYY", "COLLECTION_CODE", "CUST_CODE", "CUST_STORE_NO"}
                        .Item(C) = row.Item(C)
                    Next
                    .Item("ITEM_CATGY_CODE") = "E"
                    For I As Integer = 1 To 12
                        Dim C As String = "BUDGET_P" & Format(I, "00")
                        .Item(C) = row.Item(C)
                    Next
                End With

                If rowRSTBUDR1.Item("OPS_YYYY") <> OPS_YYYY Then
                    Dim rowbad As DataRow = dtbad.NewRow
                    rowbad.ItemArray = row.ItemArray
                    rowbad.Item("ERROR") = "Incorrect Budget Year"
                    dtbad.Rows.Add(rowbad)

                Else
                    dst.Tables("RSTBUDR1").Rows.Add(rowRSTBUDR1)
                End If

            Catch ex As Exception
                Dim rowbad As DataRow = dtbad.NewRow
                rowbad.ItemArray = row.ItemArray
                rowbad.Item("ERROR") = ex.Message
                dtbad.Rows.Add(rowbad)
            End Try
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        If dtbad.Rows.Count > 0 Then
            Using fr As New ASFMSGBF
                fr.Show_grd(dtbad, Me, "Some Rows Failed to Update - Please Check Last Column for Messages")
            End Using
        End If

    End Function

End Class