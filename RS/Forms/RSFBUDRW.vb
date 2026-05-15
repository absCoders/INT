Public Class RSFBUDRW
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim BRAND_CODE As String
    Dim RSTBUDRX As String
    Dim WKS As Integer
    Dim YWs() As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)

            ASCMAIN1.sql = "Select *" & vbCrLf _
                & " from ICTCOLL1 where BRAND_CODE = :PARM1 or :PARM1 is Null" & vbCrLf _
                & " order by COLLECTION_CODE"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, CUST_STORE_CLASS_CODE, ARTCUST2.CUST_STORE_NAME, " _
                & " NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _
                & ", CUST_STORE_CITY, CUST_STORE_STATE" _
                & " from ARTCUST2 where CUST_CODE = :PARM1 or :PARM1 is Null"
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "RSTBUDRW", "*")
            Create_BAs("RSTBUDRW")

            Create_Work_Tables()

            ASCMAIN1.sql = "Select * from " & RSTBUDRX
            Create_TDA(.Tables.Add, "RSTBUDRX", "**", 0, False, "", 3)

            Create_Relation("ARTCUSTX", "RSTBUDRX", "CUST_CODE,CUST_STORE_NO")
            With .Tables("RSTBUDRX").Columns
                Dim T As String = ""
                For W As Integer = 1 To 53
                    T &= "+ISNULL(BUDGET_P" & Format(W, "00") & ",0)"
                Next
                .Add("TOTAL", GetType(System.Decimal), Mid(T, 2))
                .Add("CUST_STORE_LOCATION", GetType(System.String), "PARENT.CUST_STORE_LOCATION")
                .Add("CUST_STORE_CLASS_CODE", GetType(System.String), "PARENT.CUST_STORE_CLASS_CODE")
                .Add("CUST_STORE_CITY", GetType(System.String), "PARENT.CUST_STORE_CITY")
                .Add("CUST_STORE_STATE", GetType(System.String), "PARENT.CUST_STORE_STATE")
            End With
        End With

        grdRSFBUDRX.DataSource = dst.Tables("RSTBUDRX")

        Create_Summary(grdRSFBUDRX, "CUST_STORE_NO", "Count")
        For M As Integer = 1 To 53
            Create_Summary(grdRSFBUDRX, "BUDGET_P" & Format(M, "00"), , , "###,##0")
        Next
        Create_Summary(grdRSFBUDRX, "TOTAL", , , "###,##0")

        'Show_Filter(grdRSTBUDRX)
        With grdRSFBUDRX.DisplayLayout.Bands("RSTBUDRX")
            For Each COLUMN_NAME As String In New String() _
                {"COLLECTION_CODE", "CUST_CODE", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For i As Integer = 1 To 53
                Dim COLUMN_NAME As String = "BUDGET_P" & Format(i, "00")
                .Columns(COLUMN_NAME).Format = "#,##0"
                .Columns(COLUMN_NAME).Header.VisiblePosition = .Columns("CUST_STORE_NO").Header.VisiblePosition + i
            Next
            .Columns("TOTAL").Format = "#,##0"

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                If New String() {"CUST_CODE", "CUST_STORE_NO", "COLLECTION_CODE", "TOTAL", "CUST_STORE_LOCATION", "CUST_STORE_CLASS_CODE", "CUST_STORE_CITY", "CUST_STORE_STATE"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    If gcol.Key <> "TOTAL" Then gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
        End With

        Dim YY As String = Mid(ASCMAIN1.CYM, 3, 2)
        If Mid(ASCMAIN1.CYM, 5, 2) >= "02" And Mid(ASCMAIN1.CYM, 5, 2) <= "07" Then
            cmdS0.Text = "S" & YY
            cmdS1.Text = "F" & YY
            cmdS2.Text = "S" & Format(Val(YY) + 1, "00")
        Else
            If Mid(ASCMAIN1.CYM, 5, 2) = "01" Then YY = Format(Val(YY) - 1, "00")
            cmdS0.Text = "F" & YY
            cmdS1.Text = "S" & Format(Val(YY) + 1, "00")
            cmdS2.Text = "F" & Format(Val(YY) + 1, "00")
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If WKS < 1 Or WKS > 53 Then
                    EMsg &= vbCr & "Weeks Range must span between 1 and 53"
                Else
                    Validate_Code("CUST_CODE", , True)
                    Validate_Code("BRAND_CODE", , True)
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("RSTBUDRW", "*") Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDRX"), New String() {"CUST_CODE", "CUST_STORE_NO"}).Select("")
                    Dim CUST_CODE As String = row.Item(0)
                    Dim CUST_STORE_NO As String = row.Item(1)
                    Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                    If rowARTCUSTX Is Nothing Then
                        EMsg &= vbCr & "Invalid Store (" & CUST_CODE & "," & CUST_STORE_NO & ")"
                    Else
                        If CUST_CODE <> "" Then
                            If rowARTCUSTX.Item("CUST_CODE") <> CUST_CODE Then
                                EMsg &= vbCr & "Store (" & CUST_CODE & "," & CUST_STORE_NO & ") does not belong to Customer " & CUST_CODE
                            End If
                        End If
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDRX"), New String() {"COLLECTION_CODE"}).Select("")
                    Dim COLLECTION_CODE As String = row.Item(0)
                    Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE})
                    If rowICTCOLL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                    Else
                        If BRAND_CODE <> "" Then
                            If rowICTCOLL1.Item("BRAND_CODE") & "" <> BRAND_CODE Then
                                EMsg &= vbCr & "Collection (" & COLLECTION_CODE & ") does not belong to Brand " & BRAND_CODE
                            End If
                        End If
                    End If
                Next

                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDRX"), New String() {"CUST_STORE_CLASS_CODE"}).Select("")
                    Dim CUST_STORE_CLASS_CODE As String = row.Item(0) & ""
                    If CUST_STORE_CLASS_CODE = "" Then
                        EMsg &= vbCr & "Some Stores Selected do NOT have a Store Class Code defined"
                    Else
                        Dim rowARTCLAS2 As DataRow = LookUp("ARTCLAS2", CUST_STORE_CLASS_CODE)
                        If rowARTCLAS2 Is Nothing Then
                            EMsg &= vbCr & "Invalid Store Class (" & CUST_STORE_CLASS_CODE & ")"
                        End If
                    End If
                Next

            Case "Load Stores"
                If optCOLLECTION_CODE.Value = "A" Then
                    EMsg &= vbCr & "Cannot Load Stores unless an Individual Collection is chosen"
                End If

            Case "Import from XLS"

                If MsgBox("This function will Import Retail Sales Budget data" _
                & vbCrLf & " from a specifically formatted spreadsheet" _
                & vbCrLf & " and use that data to replace the data currently on file" _
                & vbCrLf _
                & vbCrLf & " for the Period Range from " & Absx1.cmbFor("RYW0").Text & " thru " & Absx1.cmbFor("RYW1").Text _
                & vbCrLf _
                & vbCrLf & IIf(BRAND_CODE = "", " for All Brands", " for Brand " & BRAND_CODE & " only") _
                & vbCrLf & IIf(CUST_CODE = "", " for All Customers", " for Customer " & CUST_CODE & " only") _
                & vbCrLf _
                & vbCrLf & "Once you click 'Yes' to proceed," _
                & vbCrLf & " you will be asked for the location of the spreadsheet, " _
                & vbCrLf & " and the data will be imported and displayed in the grid below." _
                & vbCrLf _
                & vbCrLf & "You will have an opportunity to review it before clicking 'Update'." _
                & vbCrLf _
                & vbCrLf & "Proceed with the Import?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
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

                Dim COLLECTION_CODE As String = cbeCOLLECTION_CODE.Value & ""
                If COLLECTION_CODE = "" Then
                    MsgBox("Cannot Load Stores unless a Single Collection is Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                Dim sqlw As String = ""
                If CUST_CODE <> "" Then sqlw = " and CUST_CODE = '" & CUST_CODE & "'"
                For Each row As DataRow In dst.Tables("ARTCUSTX").Select(Mid(sqlw, 6))
                    Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                    Dim rowX As DataRow = dst.Tables("RSTBUDRX").Rows.Find(New Object() _
                    {COLLECTION_CODE, CUST_CODE, CUST_STORE_NO})
                    If row Is Nothing Then
                        rowX = dst.Tables("RSTBUDRX").NewRow
                        rowX.Item("COLLECTION_CODE") = COLLECTION_CODE
                        rowX.Item("CUST_CODE") = CUST_CODE
                        rowX.Item("CUST_STORE_NO") = CUST_STORE_NO
                        dst.Tables("RSTBUDRX").Rows.Add(rowX)
                    End If
                Next

            Case "Import from XLS"

                With grdRSFBUDRX.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() {"COLLECTION_CODE", "CUST_CODE", "CUST_STORE_NO"}
                        .Columns(COLUMN_NAME).Hidden = False
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                End With

                Excel_Import_SG(grdRSFBUDRX)
                Sort_grdColumns(grdRSFBUDRX, "COLLECTION_CODE,CUST_CODE,CUST_STORE_NO")
                Setup_grd()
                optCOLLECTION_CODE.CheckedIndex = 0
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
                    .Items("Load Stores").Visible = False
                    .Items("Import from XLS").Visible = ScreenMode '  Not tf
                    .Items("Show All").Visible = False ' Not tf
                End With
                .Groups("Display Options").Visible = ScreenMode And (EntryMode = "L")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        grdRSFBUDRX.Visible = ScreenMode

        With grdRSFBUDRX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = ScreenMode And (EntryMode = "L")
            .Columns("COLLECTION_CODE").Hidden = ScreenMode And (EntryMode = "L")
        End With

        If ScreenMode Then
            Setup_grd()
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUSTX", "RSTBUDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        BRAND_CODE = ""
   
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("BRAND_CODE").Text = ""
   
        If Mid(ASCMAIN1.CYM, 5, 2) >= "01" And Mid(ASCMAIN1.CYM, 5, 2) <= "06" Then
            Set_Period("S" & Mid(ASCMAIN1.CYM, 3, 2))
        Else
            Set_Period("F" & Mid(ASCMAIN1.CYM, 3, 2))
        End If

        optCOLLECTION_CODE.CheckedIndex = 0
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        BRAND_CODE = HFs("BRAND_CODE")
   
        EnforceConstraints(False)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        Fill_Records("ARTCUSTX", CUST_CODE)

        Create_Work_Tables()
        Fill_Records("RSTBUDRX")

        Fill_Records("ICTCOLL1", BRAND_CODE)

        EnforceConstraints(True)

        Sort_grdColumns(grdRSFBUDRX, "CUST_STORE_NO")

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & " from ICTCOLL1,ICTBRAN1" & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & IIf(BRAND_CODE = "", "", " and ICTCOLL1.BRAND_CODE = '" & BRAND_CODE & "'" & vbCrLf) _
            & " order by ICTCOLL1.COLLECTION_CODE"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        'cbeCOLLECTION_CODE.DataSource = dst.Tables("ICTCOLL1")
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Set_Column_Headings()

        Setup_grd()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim RYW0 As String = Absx1.cmbFor("RYW0").Value
        Dim RYW1 As String = Absx1.cmbFor("RYW1").Value

        Dim sql_Delete As String = "Delete from RSTBUDRW" _
            & " where OPS_YYYYWW between '" & RYW0 & "' and '" & RYW1 & "'" _
            & IIf(BRAND_CODE = "", "", " and COLLECTION_CODE in (Select COLLECTION_CODE from ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "')" & vbCrLf) _
            & IIf(CUST_CODE = "", "", " and CUST_CODE = '" & CUST_CODE & "'" & vbCrLf)


        dst.Tables("RSTBUDRW").Rows.Clear()

        For Each rowRSTBUDRX As DataRow In dst.Tables("RSTBUDRX").Select("")
            For I As Integer = 1 To WKS
                Dim BUDGET As Decimal = Val(rowRSTBUDRX.Item("BUDGET_P" & Format(I, "00")) & "")
                If BUDGET <> 0 Then
                    Dim rowRSTBUDRW As DataRow = dst.Tables("RSTBUDRW").NewRow
                    rowRSTBUDRW.Item("COLLECTION_CODE") = rowRSTBUDRX.Item("COLLECTION_CODE")
                    rowRSTBUDRW.Item("CUST_CODE") = rowRSTBUDRX.Item("CUST_CODE")
                    rowRSTBUDRW.Item("CUST_STORE_NO") = rowRSTBUDRX.Item("CUST_STORE_NO")
                    rowRSTBUDRW.Item("OPS_YYYYWW") = YWs(I)
                    rowRSTBUDRW.Item("BUDGET") = BUDGET
                    dst.Tables("RSTBUDRW").Rows.Add(rowRSTBUDRW)
                End If
            Next
        Next

        ' Update_Record_TDA("RSTBUDRW", sql_Delete)

        ASCMAIN1.sql = sql_Delete
        ASCDATA1.ExecuteSQL()
        Update_BAs("RSTBUDRW")

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
        Load_Popup_Menu(grdRSFBUDRX, "SS", "Show Filter", "Show GroupBox")
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
                Case "grdRSTBUDRX"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Clear Column"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")
                    'tlb_btn = DirectCast(tlb_pop.Tools("Copy Value"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = ScreenMode And (EntryMode = "L")

                    If grdRSFBUDRX.Tag = "" Then
                        'e.Cancel = True
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Clear Column"
            '    Dim COLUMN_NAME As String = grdRSTBUDRX.Tag
            '    If COLUMN_NAME = "" Then Exit Sub
            '    If COLUMN_NAME = "CUST_STORE_NO" Then Exit Sub
            '    For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            '        row.Item(COLUMN_NAME) = DBNull.Value
            '    Next
            'Case "Copy Value"
            '    Dim COLUMN_NAME As String = grdRSTBUDRX.Tag
            '    If COLUMN_NAME = "" Then Exit Sub
            '    If grdRSTBUDRX.ActiveRow Is Nothing OrElse grdRSTBUDRX.ActiveRow.IsAddRow OrElse Not grdRSTBUDRX.ActiveRow.IsDataRow Then Exit Sub
            '    Dim COPY_VALUE As String = grdRSTBUDRX.ActiveRow.Cells(COLUMN_NAME).Value
            '    For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.CurrentRows)
            '        row.Item(COLUMN_NAME) = COPY_VALUE
            '    Next

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
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

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If Absx1.cmbFor("RYW0").Value & "" <> "" And Absx1.cmbFor("RYW1").Value & "" <> "" Then
            WKS = 1 + ASCMAIN1.Week_Diff(Absx1.cmbFor("RYW0").Value, Absx1.cmbFor("RYW1").Value)
            lblWeeks.Text = CStr(WKS) & " Wks"
        Else
            lblWeeks.Text = ""
        End If

        Dim RYW0 As String = Absx1.cmbFor("RYW0").Value
        If WKS >= 1 And WKS <= 53 Then
            ReDim YWs(WKS)
            For i As Integer = 1 To WKS
                YWs(i) = ASCMAIN1.Week_Calc(RYW0, (i - 1))
            Next
        End If

    End Sub
#End Region

#Region "grdRSTBUDRX"

    Private Sub grdRSTBUDRX_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSFBUDRX.AfterExitEditMode
        With grdRSFBUDRX
            Select Case .ActiveCell.Column.Key
                Case "COLLECTION_CODE"
                    Dim COLLECTION_CODE As String = .ActiveCell.Text
                    If COLLECTION_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(COLLECTION_CODE, .ActiveCell.Column.Key)
                    End If

                Case "CUST_CODE"
                    Dim CUST_CODE As String = .ActiveCell.Text
                    If CUST_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(CUST_CODE, .ActiveCell.Column.Key)
                    End If

                Case "CUST_STORE_NO"
                    Dim CUST_STORE_NO As String = .ActiveCell.Text
                    If CUST_STORE_NO <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(CUST_STORE_NO, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTBUDRX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSFBUDRX.AfterRowActivate
        With grdRSFBUDRX.DisplayLayout.Bands(0)
            If grdRSFBUDRX.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdRSTBUDRX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSFBUDRX.BeforeRowUpdate
        With grdRSFBUDRX

            If Not e.Cancel Then
                If e.Row.Cells("CUST_CODE").Text = "" And ScreenMode Then
                    .ActiveRow.Cells("CUST_CODE").Value = Absx1.CtlFor("CUST_CODE").Text
                End If
                If e.Row.Cells("COLLECTION_CODE").Text = "" And ScreenMode Then
                    .ActiveRow.Cells("COLLECTION_CODE").Value = cbeCOLLECTION_CODE.Value
                End If

            End If
        End With
    End Sub

    Private Sub grdRSTBUDRX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSFBUDRX.ClickCellButton
        Select Case grdRSFBUDRX.ActiveCell.Column.Key
            Case "COLLECTION_CODE"
                grdClickCellButton(grdRSFBUDRX)
            Case "CUST_CODE"
                grdClickCellButton(grdRSFBUDRX)
            Case "CUST_STORE_NO"
                'grdClickCellButton(grdRSTBUDRX, "CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Text & "'")
                grdClickCellButton(grdRSFBUDRX, "CUST_CODE = '" & HFs("CUST_CODE") & "'")
        End Select
    End Sub
#End Region

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grd()
    End Sub

    Sub Setup_grd()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdRSTBUDRX.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdRSFBUDRX.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").Hidden = (optCOLLECTION_CODE.Value <> "A")
            .Columns("CUST_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")

            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdRSTBUDRX.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("CUST_STORE_NO", False)
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdRSFBUDRX.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            allow_modifications = False
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        Dim RYP_LEGENDS As String = Absx1.cmbFor("RYW0").Text & " thru " & Absx1.cmbFor("RYW1").Text

        DVW.RowFilter = Mid(sql, 5)
        grdRSFBUDRX.Text = "Retail Sales Budgets, by Store/Week, for " & RYP_LEGENDS & " - " & COLLS
        If allow_modifications Then
            grdRSFBUDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdRSFBUDRX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdRSFBUDRX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        Else
            grdRSFBUDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdRSFBUDRX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

            'grdRSTBUDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            'grdRSTBUDRX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdRSFBUDRX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        End If
        'grdRSTBUDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grd()
    End Sub

    Sub Set_Column_Headings()

        Dim RYW0 As String = Absx1.cmbFor("RYW0").Value
        Dim RYW1 As String = Absx1.cmbFor("RYW1").Value

        For M As Integer = 1 To 53
            With grdRSFBUDRX.DisplayLayout.Bands(0).Columns("BUDGET_P" & Format(M, "00"))
                Dim YP As String = ASCMAIN1.Week_Calc(RYW0, (M - 1))
                If YP > RYW1 Then
                    .Hidden = True
                Else
                    Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", YP)
                    Dim LEGEND As String = rowGLTPARM3.Item("LEGEND")
                    .Header.Caption = Mid(LEGEND, 10, 6)
                    .Width = 60
                    .Hidden = False
                End If
            End With
        Next
    End Sub

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

        load_handled = True
        If dt.Rows.Count = 0 Then
            MsgBox("No Rows Loaded", MsgBoxStyle.OkOnly, "Import Failed")
        Else
            dst.Tables("RSTBUDRX").Rows.Clear()
        End If

        For Each row As DataRow In dt.Select("")
            r += 1
            If r Mod 100 = 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Budgets from XLS")
                RowsMax = dt.Rows.Count
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RowsMax))
            End If

            Try
                Dim rowRSTBUDRX As DataRow = dst.Tables("RSTBUDRX").NewRow
                With rowRSTBUDRX
                    For Each C As String In New String() {"COLLECTION_CODE", "CUST_CODE", "CUST_STORE_NO"}
                        .Item(C) = row.Item(C)
                    Next
                    For I As Integer = 1 To WKS
                        Dim C As String = "BUDGET_P" & Format(I, "00")
                        .Item(C) = row.Item(C)
                    Next
                End With

                dst.Tables("RSTBUDRX").Rows.Add(rowRSTBUDRX)

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

    Private Sub cmdS0_Click(sender As System.Object, e As System.EventArgs) Handles cmdS0.Click
        Set_Period(cmdS0.Text)
    End Sub

    Private Sub cmdS1_Click(sender As System.Object, e As System.EventArgs) Handles cmdS1.Click
        Set_Period(cmdS1.Text)
    End Sub

    Private Sub cmdS2_Click(sender As System.Object, e As System.EventArgs) Handles cmdS2.Click
        Set_Period(cmdS2.Text)
    End Sub

    Sub Set_Period(SYY As String)
        Dim YY As String = Mid(SYY, 2, 2)

        Dim P0 As String = ""
        Dim P1 As String = ""

        If Mid(SYY, 1, 1) = "S" Then
            P0 = "20" & YY & "01"
            P1 = "20" & YY & "26"
        Else
            P0 = "20" & YY & "27"
            P1 = "20" & YY & "52"
        End If

        If optCalendar.Value = "O" Then
            P0 = ASCMAIN1.Week_Calc(P0, -4)
            P1 = ASCMAIN1.Week_Calc(P1, -4)
        End If

        ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYWW = '" & P0 & "'"
        Absx1.cmbFor("RYW0").Value = ASCDATA1.GetDataValue
        ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYWW = '" & P1 & "'"
        Absx1.cmbFor("RYW1").Value = ASCDATA1.GetDataValue
    End Sub

    Private Sub optCalendar_ValueChanged_1(sender As Object, e As EventArgs) Handles optCalendar.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Dim YW As String

        YW = Absx1.cmbFor("RYW0").Value
        If optCalendar.Value = "R" Then
            YW = ASCMAIN1.Week_Calc(YW, 1)
        Else
            YW = ASCMAIN1.Week_Calc(YW, -1)
        End If
        Absx1.cmbFor("RYW0").Value = YW

        YW = Absx1.cmbFor("RYW1").Value
        If optCalendar.Value = "R" Then
            YW = ASCMAIN1.Week_Calc(YW, 1)
        Else
            YW = ASCMAIN1.Week_Calc(YW, -1)
        End If
        Absx1.cmbFor("RYW1").Value = YW

    End Sub

    Sub Create_Work_Tables()

        Dim RYW0 As String = Absx1.cmbFor("RYW0").Value
        Dim RYW1 As String = Absx1.cmbFor("RYW1").Value
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text

        Dim sqlM As String = ""
        For I As Integer = 1 To 53
            Dim YW As String = ASCMAIN1.Week_Calc(RYW0, I - 1)
            If YW > RYW1 Then YW = ""
            sqlM &= ", Sum (Decode(RSTBUDRW.OPS_YYYYWW,'" & YW & "',BUDGET,0)) BUDGET_P" & Format(I, "00")
        Next

        ASCMAIN1.sql = "Select RSTBUDRW.COLLECTION_CODE, RSTBUDRW.CUST_CODE, RSTBUDRW.CUST_STORE_NO" & vbCrLf _
            & sqlM _
            & " from RSTBUDRW,ARTCUST2" _
            & " where RSTBUDRW.OPS_YYYYWW between '" & RYW0 & "' and '" & RYW1 & "'" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = RSTBUDRW.CUST_CODE and ARTCUST2.CUST_STORE_NO = RSTBUDRW.CUST_STORE_NO" & vbCrLf _
            & IIf(CUST_CODE = "", "", "   and RSTBUDRW.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf) _
            & IIf(BRAND_CODE = "", "", "" _
                  & " and RSTBUDRW.COLLECTION_CODE in " _
                  & " (Select COLLECTION_CODE FROM ICTCOLL1 where BRAND_CODE = '" & BRAND_CODE & "')" & vbCrLf) _
            & " group by RSTBUDRW.COLLECTION_CODE, RSTBUDRW.CUST_CODE, RSTBUDRW.CUST_STORE_NO"

        If RSTBUDRX = "" Then
            RSTBUDRX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTBUDRX & " Add Primary Key (COLLECTION_CODE, CUST_CODE, CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTBUDRX)
            ASCDATA1.ExecuteSQL("Insert into " & RSTBUDRX & " " & ASCMAIN1.sql)
        End If
    End Sub
End Class