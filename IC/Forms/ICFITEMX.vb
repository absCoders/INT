Public Class ICFITEMX

    Dim sqlICTITEMX As String
    Dim ICTITEMX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst

            ASCMAIN1.sql = "Select * from ICTITEMX where ROWNUM < 1"
            ICTITEMX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add Primary Key (ITEM_CODE)")

            ASCMAIN1.sql = "Select ICTITEMX.*" _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.COLLECTION_CODE" _
            & " from ICTITEM1," & ICTITEMX & " ICTITEMX" _
            & " where ICTITEM1.ITEM_CODE = ICTITEMX.ITEM_CODE"
            sqlICTITEMX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, ICTITEMX, "**", 0)
            .Tables(ICTITEMX).Columns.Add("BAD_ITEM")
        End With

        grdICTITEMX.DataSource = dst.Tables(ICTITEMX)

        With grdICTITEMX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key <> "ITEM_CATGY_CODE" Then
                '    gcol.CellAppearance.BackColor = Color.Beige
                'End If
                If gcol.Key = "ITEM_DESC" _
                Or gcol.Key = "ITEM_CATGY_CODE" _
                Or gcol.Key = "COLLECTION_CODE" _
                Or gcol.Key = "BAD_ITEM" Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
            '.Columns("ITEM_CATGY_CODE").CellAppearance.BackColor = Color.Beige
        End With

        Create_Summary(grdICTITEMX, "ITEM_CODE", "Count")

        ASCMAIN1.Add_Value_List(grdICTITEMX, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMX, "MATL_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMX, "STONE_CLASS_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMX, "METAL_CLASS_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEMX, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTITEMX_D", optData.Value) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                Dim COL_TBL As New Dictionary(Of String, String)
                COL_TBL.Add("CUST_CODE", "ARTCUST1")
                COL_TBL.Add("MATL_CATGY_CODE", "ICTLATMA")
                COL_TBL.Add("STONE_CLASS_CODE", "ICTMATLB")
                COL_TBL.Add("METAL_CLASS_CODE", "ICTMATLC")

                For Each COLUMN_NAME In COL_TBL.Keys
                    Dim BAD_CODES As String = ""
                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables(ICTITEMX).Select, New String() {COLUMN_NAME}).Rows
                        Dim CODE_VALUE As String = row.Item(0) & ""
                        If CODE_VALUE <> "" Then
                            If LookUp(COL_TBL(COLUMN_NAME), CODE_VALUE) Is Nothing Then
                                BAD_CODES &= "," & CODE_VALUE
                            End If
                        End If
                    Next
                    If BAD_CODES <> "" Then
                        MsgBox("Please correct the following codes:" _
                        & vbCrLf & vbCrLf & Mid(BAD_CODES, 2), _
                        MsgBoxStyle.OkOnly, "Bad Code Values found for " & COLUMN_NAME)
                        EMsg &= vbCr & "Bad Code Values found for " & COLUMN_NAME & "; please correct before Updating"
                    End If
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Load from Spreadsheet"
                Load_from_Spreadsheet()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load from Spreadsheet").Settings.Enabled = iScreenMode

                .Groups("Upload from Spreadsheet").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        With grdICTITEMX.DisplayLayout.Bands(0)
            .Columns("ITEM_SAFETY_STOCK").Hidden = Not (optData.Value = "Q")
            .Columns("CUST_CODE").Hidden = Not (optData.Value = "X")
            .Columns("LAUNCH_DATE").Hidden = Not (optData.Value = "X")
            .Columns("MATL_CATGY_CODE").Hidden = Not (optData.Value = "C")
            .Columns("STONE_CLASS_CODE").Hidden = Not (optData.Value = "C")
            .Columns("METAL_CLASS_CODE").Hidden = Not (optData.Value = "C")
        End With

        optData.Enabled = Not ScreenMode

        grdICTITEMX.Visible = ScreenMode
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables(ICTITEMX).Rows.Clear()
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        Dim sqlw = ""
        Select Case optData.Value
            Case "Q"
                sqlw = " where ICTITEMX.ITEM_SAFETY_STOCK <> 0"
            Case "X"
                sqlw = " where (ICTITEMX.CUST_CODE IS NOT NULL OR ICTITEMX.LAUNCH_DATE IS NOT NULL)"
            Case "C"
                sqlw = " where (ICTITEMX.MATL_CATGY_CODE IS NOT NULL OR ICTITEMX.STONE_CLASS_CODE IS NOT NULL OR ICTITEMX.METAL_CLASS_CODE IS NOT NULL)"
        End Select

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMX)
        ASCDATA1.ExecuteSQL("Insert into " & ICTITEMX & " Select * from ICTITEMX " & sqlw)

        Fill_Records(ICTITEMX)
        EnforceConstraints(True)

        Sort_grdColumns(grdICTITEMX, "ITEM_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables(ICTITEMX).AcceptChanges()

        Update_Record_TDA(ICTITEMX, "1=1")

        Dim sqlSet As String = ""
        Dim sqlInsert As String = ""
        Dim sqlColumns As String = ""

        'Select Case optData.Value
        '    Case "Q"
        '        sqlSet = "Set ITEM_SAFETY_STOCK = R1.ITEM_SAFETY_STOCK"
        '        sqlColumns = "R1.ITEM_SAFETY_STOCK"
        '    Case "X"
        '        sqlSet = "Set CUST_CODE = R1.CUST_CODE, LAUNCH_DATE = R1.LAUNCH_DATE"
        '        sqlColumns = "R1.ITEM_CODE,R1.ITEM_SAFETY_STOCK,R1.LAUNCH_DATE"
        '    Case "C"
        '        sqlSet = "Set MATL_CATGY_CODE = R1.MATL_CATGY_CODE, STONE_CLASS_CODE = R1.STONE_CLASS_CODE, METAL_CLASS_CODE = R1.METAL_CLASS_CODE"
        '        sqlColumns = "R1.ITEM_CODE,R1.MATL_CATGY_CODE,R1.STONE_CLASS_CODE,R1.METAL_CLASS_CODE"
        'End Select

        Dim COLUMN_NAMEs As New List(Of String)
        Select Case optData.Value
            Case "Q"
                COLUMN_NAMEs.Add("ITEM_SAFETY_STOCK")
                'sqlSet = "Set ITEM_SAFETY_STOCK = R1.ITEM_SAFETY_STOCK"
                'sqlColumns = "R1.ITEM_SAFETY_STOCK"
            Case "X"
                COLUMN_NAMEs.Add("CUST_CODE")
                COLUMN_NAMEs.Add("LAUNCH_DATE")
                'sqlSet = "Set CUST_CODE = R1.CUST_CODE, LAUNCH_DATE = R1.LAUNCH_DATE"
                'sqlColumns = "R1.ITEM_CODE,R1.ITEM_SAFETY_STOCK,R1.LAUNCH_DATE"
            Case "C"
                COLUMN_NAMEs.Add("MATL_CATGY_CODE")
                COLUMN_NAMEs.Add("STONE_CLASS_CODE")
                COLUMN_NAMEs.Add("METAL_CLASS_CODE")
                'sqlSet = "Set MATL_CATGY_CODE = R1.MATL_CATGY_CODE, STONE_CLASS_CODE = R1.STONE_CLASS_CODE, METAL_CLASS_CODE = R1.METAL_CLASS_CODE"
                'sqlColumns = "R1.ITEM_CODE,R1.MATL_CATGY_CODE,R1.STONE_CLASS_CODE,R1.METAL_CLASS_CODE"
        End Select

        Dim sqlNull As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            sqlSet &= ", " & COLUMN_NAME & " = R1." & COLUMN_NAME
            sqlColumns &= ",R1." & COLUMN_NAME
            sqlNull &= "," & COLUMN_NAME & " = Null"
        Next
        sqlSet = "Set " & Mid(sqlSet, 3)
        sqlColumns = Mid(sqlColumns, 2)

        ASCDATA1.ExecuteSQL("Update ICTITEMX Set " & Mid(sqlNull, 2))

        sqlInsert = "(" & Replace(sqlColumns, "R1.", "") & ") Values (" & sqlColumns & ")"

        ASCMAIN1.sql = "" _
        & "Begin " _
        & " Declare Cursor C1 is Select * from " & ICTITEMX & ";" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update ICTITEMX " _
        & "    " & sqlSet _
        & "    where ITEM_CODE = R1.ITEM_CODE;" _
        & "   If SQL%NOTFOUND Then " _
        & "    Insert into ICTITEMX " & sqlInsert & ";" _
        & "   End If;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & "BEGIN DECLARE CURSOR C1 IS SELECT * FROM ICTITEMX;" _
        & " BEGIN " _
        & " UPDATE ICTITEM1 SET " _
        & "  CUST_CODE = NULL" _
        & ", LAUNCH_DATE = NULL" _
        & ", ITEM_SAFETY_STOCK = NULL" _
        & ", MATL_CATGY_CODE = NULL" _
        & ", STONE_CLASS_CODE = NULL" _
        & ", METAL_CLASS_CODE = NULL;" _
        & " FOR R1 IN C1 LOOP" _
        & " UPDATE ICTITEM1 SET " _
        & "  CUST_CODE = R1.CUST_CODE" _
        & ", LAUNCH_DATE = R1.LAUNCH_DATE" _
        & ", ITEM_SAFETY_STOCK = R1.ITEM_SAFETY_STOCK" _
        & ", MATL_CATGY_CODE = R1.MATL_CATGY_CODE" _
        & ", STONE_CLASS_CODE = R1.STONE_CLASS_CODE" _
        & ", METAL_CLASS_CODE = R1.METAL_CLASS_CODE" _
        & " WHERE ITEM_CODE = R1.ITEM_CODE;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTITEMX, "SSSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Bad Items Only", "Load from Spreadsheet")
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

                'Case "grdDPTFCSTD"
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Load from Spreadsheet"
                Load_from_Spreadsheet()

            Case "Show Bad Items Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim dvw As DataView = DirectCast(grdICTITEMX.DataSource, DataTable).DefaultView
                Dim grdCaption As String = "Extended Item Master Data (" & optData.Text & ")"
                If tlb_sbt.Checked Then
                    dvw.RowFilter = "ISNULL(BAD_ITEM,'0') = '1'"
                    'ASCMAIN1.Notify("Now Showing Only those Items with Differences")
                    grdICTITEMX.Text = grdCaption & " - Bad Items Only"
                Else
                    dvw.RowFilter = ""
                    'ASCMAIN1.Notify("Now Showing All Items")
                    grdICTITEMX.Text = grdCaption
                End If

        End Select

        Select Case grd.Name
            'Case "grdDPTFCSTD"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
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
            'Case "BRAND_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        'Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "BRAND_CODE"
            '    'Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            'Case "BRAND_CODE"
            '    If EntryMode = "" Then
            '        If Absx1.txtFor("BRAND_CODE").Text <> "" Then
            '            Call LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
            '            If cdr IsNot Nothing Then

            '            End If
            '        End If
            '    End If

        End Select
    End Sub

#End Region

#Region "Excel Upload"

    Sub Load_from_Spreadsheet()

        Dim tbl As DataTable = dst.Tables(ICTITEMX).Clone
        tbl.Merge(dst.Tables(ICTITEMX))

        dst.Tables(ICTITEMX).Rows.Clear()

        If Excel_Import(grdICTITEMX) = -1 Then
            dst.Tables(ICTITEMX).Merge(tbl)
        Else
            If optReplace.Value = "I" Then
                For Each row As DataRow In tbl.Rows
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                    Dim rowICTITEMX As DataRow = dst.Tables(ICTITEMX).Rows.Find(ITEM_CODE)
                    If rowICTITEMX Is Nothing Then
                        dst.Tables(ICTITEMX).Rows.Add(row.ItemArray)
                    End If
                Next
            End If

            Sort_grdColumns(grdICTITEMX, "ITEM_CODE")
            grdICTITEMX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        End If

        If optReplace.Value = "A" Then
            If MsgBox("You have selected to Replace All " & CStr(tbl.Select.Length) _
                      & " items previously displayed" _
                      & vbCrLf & " with the " & dst.Tables(ICTITEMX).Select.Length _
                      & " items currently on display." _
                      & vbCrLf & vbCrLf & "OK to Continue?", _
                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                dst.Tables(ICTITEMX).Rows.Clear()
                dst.Tables(ICTITEMX).Merge(tbl)
            End If
        End If

    End Sub

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid, _
     Optional ByRef load_by_table As Boolean = False, _
     Optional ByRef load_handled As Boolean = False, _
     Optional ByRef F As ASFEXCL1 = Nothing)

        If grd.Name = "grdICTITEMX" Then
            load_by_table = True
        End If

    End Sub

    Overrides Sub Excel_Import_Post_Process(ByVal grd As UltraWinGrid.UltraGrid, F As ASFEXCL1)

        Dim BAD_ITEM_count As Int64 = dst.Tables(ICTITEMX).Select("BAD_ITEM = '1'").Length
        If BAD_ITEM_count <> 0 Then
            MsgBox("There were " & CStr(BAD_ITEM_count) _
                   & " Bad Items Loaded from Spreadsheet", _
                   MsgBoxStyle.OkOnly, _
                   CStr(dst.Tables(ICTITEMX).Rows.Count) & " Records Loaded")
        Else
            MsgBox("All Items Loaded Successfully", MsgBoxStyle.OkOnly, CStr(dst.Tables(ICTITEMX).Rows.Count) & " Records Loaded")
        End If

    End Sub

    Overrides Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow, _
     Optional ByVal TBL As DataTable = Nothing)

        For Each COLUMN_NAME As String In New String() _
            {"CUST_CODE", "MATL_CATGY_CODE", "STONE_CLASS_CODE", "METAL_CLASS_CODE"}
            If CStr(row.Item(COLUMN_NAME) & "").Length > dst.Tables(ICTITEMX).Columns(COLUMN_NAME).MaxLength Then
                row.Item(COLUMN_NAME) = Mid(row.Item(COLUMN_NAME), 1, dst.Tables(ICTITEMX).Columns(COLUMN_NAME).MaxLength)
            End If
        Next

        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
        row.Item("ITEM_DESC") = ""
        row.Item("COLLECTION_CODE") = ""
        row.Item("ITEM_CATGY_CODE") = ""
        row.Item("BAD_ITEM") = ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 IsNot Nothing Then
            row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            row.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
            row.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
        Else
            row.Item("BAD_ITEM") = "1"
        End If
    End Sub
#End Region

#Region "grdICTITEMX"

    Private Sub grdICTITEMX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMX.AfterCellUpdate

        If e.Cell.Column.Key = "ITEM_CODE" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", e.Cell.Text)
            If rowICTITEM1 IsNot Nothing Then
                e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                e.Cell.Row.Cells("ITEM_CATGY_CODE").Value = rowICTITEM1.Item("ITEM_CATGY_CODE") & ""
                e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE") & ""
                e.Cell.Row.Cells("BAD_ITEM").Value = ""
            Else
                e.Cell.Row.Cells("ITEM_DESC").Value = ""
                e.Cell.Row.Cells("BAD_ITEM").Value = "1"
            End If
        End If

    End Sub

    Private Sub grdICTITEMX_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTITEMX.BeforeExitEditMode
        grdFieldFormat(grdICTITEMX)
    End Sub

    Private Sub grdICTITEMX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTITEMX.BeforeRowUpdate
        With grdICTITEMX
            If LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text) Is Nothing Then
                MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
                Exit Sub
            End If

            Select Case optData.Value
                Case "Q"
                    If Val(e.Row.Cells("ITEM_SAFETY_STOCK").Value & "") < 0 Then
                        MsgBox("Invalid Negative Value entered for Safety Stock (" & e.Row.Cells("ITEM_SAFETY_STOCK").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                        Exit Sub
                    End If

                Case "X"
                    If e.Row.Cells("CUST_CODE").Text <> "" Then
                        If LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text) Is Nothing Then
                            MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If


                Case "C"
                    For Each COLUMN_NAME As String In New String() {"MATL_CATGY_CODE", "STONE_CLASS_CODE", "METAL_CLASS_CODE"}
                        If e.Row.Cells(COLUMN_NAME).Text <> "" Then
                            Dim TABLE_NAME As String = IIf(COLUMN_NAME = "MATL_CATGY_CODE", "ICTMATLA", IIf(COLUMN_NAME = "STONE_CLASS_CODE", "ICTMATLB", "ICTMATLC"))
                            If LookUp(TABLE_NAME, e.Row.Cells(COLUMN_NAME).Text) Is Nothing Then
                                MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                                e.Cancel = True
                                Exit Sub
                            End If
                        End If
                    Next

            End Select
        End With
    End Sub

    Private Sub grdICTITEMX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTITEMX.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdICTITEMX, sql_where, False)
    End Sub

    Private Sub grdICTITEMX_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTITEMX.Error
        grdICTITEMX.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdICTITEMX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEMX.InitializeRow
        If e.Row.Cells("BAD_ITEM").Value & "" = "1" Then
            e.Row.Cells("BAD_ITEM").Appearance.BackColor = Color.Red
        Else
            e.Row.Cells("BAD_ITEM").Appearance.BackColor = Color.Empty
        End If
    End Sub
#End Region

End Class