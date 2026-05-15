Public Class SPFMXWS1

#Region "Declarations"
    Dim rowSPTMXWS1 As DataRow

    Dim rowSOTSELL1 As DataRow
    Dim rowICTSEAS1 As DataRow

    Dim SELL_CODE As String
    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String

    Dim SELL_CODE_this_user As String
    Dim REGION_CODE_this_user As String

    Dim SEASON_CODE_LY As String

    Dim XLSR As New Dictionary(Of String, Integer)

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim rangeCopyFrom As SpreadsheetGear.IRange
    Dim rangePaste_To As SpreadsheetGear.IRange

    Dim xls_STOREs As New List(Of String)
    Dim xls_CBs As New List(Of String)
    'Dim xls_STs As New List(Of String)
    Dim xls_STs As New Dictionary(Of String, List(Of String))
    Dim xls_NOT_STs As New List(Of String)

    Dim LOCK_SHEETS As New Dictionary(Of String, String)

    Dim XLS_NO As String
    Dim XLS_PWD As String = "ABS"

    Dim isAC As Boolean = False
    Dim load_from_AC As Boolean = False
    Dim isAE_and_Update_AC As Boolean = False

    Dim SPTMXWS2 As String
    Dim ARTCUST2 As String
    Dim RSTBUDR1 As String
    Dim SPTCWRXW As String

    Dim SPTMXWS5 As String

    Dim YPs() As String
    Dim TYWKs()
    Dim LYWKs()

    Dim AYW As String
    Dim AYWi As Integer
    Dim Number_of_Weeks_in_Season As Integer

    Dim blnConsolidated As Boolean = False
    Dim SELL_CODES_consolidated As New List(Of String)

    Dim CBi As Integer = 0  ' Number of Checkbooks
    Dim COLS_Recap As Integer = 10 + 3 + 1  ' Number of Columns per Checkbook on Recap Sheet; adding 3 for ST +1 for spacer
    Dim COLS_Summary As Integer = 5 ' Number of Columns per Checkbook on Summary Sheet

    Dim COLS_ST As Integer = 1 ' Number of Columns of Checkbook Sub-Totals, plus 1 for a spacer
    Dim colPNOTES As Integer = 1 ' Relative Column for Promo Notes (used for user entered notes in Total section)
    Dim colLYRHRS As Integer = 2 ' Relative Column for Last Year Actual Modeling Hours
    Dim colLYRACT As Integer = 3 ' Relative Column for Last Year Actual Modeling Spend
    Dim colLYRSLS As Integer = 4 ' Relative Column for Last Year Actual Retail Sales
    Dim colBUDRTL As Integer = 5 ' Relative Column for Retail Sales Budget, Actualized
    Dim colPCTRTL As Integer = 6 ' Relative Column for Retail Sales %Change from LY
    Dim colBUDHRS As Integer = 7 ' Relative Column for Modeling Hours, Actualized
    Dim colBUDMOD As Integer = 8 ' Relative Column for Modeling Budget, Actualized
    Dim colCVGMOD As Integer = 9 ' Relative Column for Modeling Coverage %

    Dim colST_SLS As Integer = 11 ' Relative Column for Sub-Total for SLS
    Dim colST_HRS As Integer = 12 ' Relative Column for Sub-Total for HRS
    Dim colST_CVG As Integer = 13 ' Relative Column for Sub-Total for Modeling Coverage %

    Dim colPLANRATE As Integer = colLYRACT ' Relative Column for Plan Rate
    Dim colHOURS2PLAN As Integer = colCVGMOD ' Relative Column for Total Hours to Plan

    ' Dim SCOL_Recap As Integer = 6 ' Number of columns before the Total Checkbook block in the Recap Sheet (ie, the Notes column)
    Dim SCOL_Recap As Integer = 5 ' 0 based Starting Column for column just before the Total Checkbook block in Recap Sheet (ie, the Notes column)
    Dim SCOL_Summary As Integer = 6 ' 0 based Starting Column for Total Checkbook block in Summary Sheet

    Dim SROW_Weeks As Integer = 17  ' 0 based Starting Row for Weeks Heading
    Dim SROW_Months As Integer = 7  ' 0 based Starting Row for Months Heading

    Dim SaveMode As Boolean = False
    Dim YW_LAST_CWRX As String = ""

    Dim SEASON_SLS_BUD_LOCKED As String
    Dim CVG_PCT() As String

    Dim useFrozenAlignment As Boolean
    Dim frozenAlignment_YP As String

    Dim automate_Initialization As Boolean = False

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.USER_CODES.Contains("FS") Then
            SELL_CODE_this_user = "?"
            Dim rowTATUSER1 As DataRow = LookUp("TATUSER1", ASCMAIN1.USER_ID)
            If rowTATUSER1 IsNot Nothing Then
                SELL_CODE_this_user = rowTATUSER1.Item("SELL_CODE") & ""
                REGION_CODE_this_user = rowTATUSER1.Item("REGION_CODE") & ""
                If SELL_CODE_this_user = "" And REGION_CODE_this_user = "" Then
                    SELL_CODE_this_user = "?"
                End If

                Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE_this_user)
                If rowSOTSELL1 IsNot Nothing Then
                    If rowSOTSELL1.Item("SELL_TYPE") & "" = "AC" Then
                        isAC = True
                    End If
                End If
            End If
        End If

        Create_ARTCUST2()

        With dst
            Create_TDA(.Tables.Add, "SOTSELL1", "*", 1)
            Create_TDA(.Tables.Add, "SPTMXWS0", "*", 1)
            Create_TDA(.Tables.Add, "SPTMXWS1", "*", 2)
            Create_TDA(.Tables.Add, "SPTMXWS2", "*", 0)
            Create_TDA(.Tables.Add, "SPTMXWS3", "*", 0)
            Create_TDA(.Tables.Add, "SPTMXWS4", "*", 0)
            Create_TDA(.Tables.Add, "SPTMXWSS", "*", 1)

            Create_TDA(.Tables.Add, "SPTMXAC2", "*", 0)
            Create_TDA(.Tables.Add, "SPTMXAC3", "*", 0)
            Create_TDA(.Tables.Add, "SPTMXAC4", "*", 0)

            Create_TDA(.Tables.Add, "SPTCWRXC", "*", 0, False)
            .Tables("SPTCWRXC").Columns.Add("CBI", GetType(System.Int16))

            ASCMAIN1.sql = "Select * from " & ARTCUST2 & " ARTCUST2"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 2)
            .Tables("ARTCUST2").Columns.Add("HAS_BUDGET")
            .Tables("ARTCUST2").Columns("HAS_BUDGET").DefaultValue = "0"

            ASCMAIN1.sql = "Select SPTMXWS1.*,SOTSELL1.SELL_NAME,SOTSELL1.REGION_CODE" & vbCrLf _
                & " from SPTMXWS1,SOTSELL1" & vbCrLf _
                & " where SOTSELL1.SELL_CODE (+) = SPTMXWS1.SELL_CODE" & vbCrLf _
                & "   and SPTMXWS1.SEASON_CODE = :PARM1"
            If SELL_CODE_this_user <> "" Then
                ASCMAIN1.sql &= " and (SPTMXWS1.SELL_CODE = '" & SELL_CODE_this_user & "' or SOTSELL1.SELL_CODE_MGR = '" & SELL_CODE_this_user & "')"
            End If
            If REGION_CODE_this_user <> "" Then
                ASCMAIN1.sql &= " and SPTMXWS1.SELL_CODE in" & vbCrLf _
                    & " (Select SELL_CODE from SOTSELL1 where REGION_CODE = '" & REGION_CODE_this_user & "')"
            End If
            Create_TDA(.Tables.Add, "SPTMXWSX", "**", 0, False, "V", 2)
            .Tables("SPTMXWSX").Columns.Add("SEL")
            .Tables("SPTMXWSX").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select * from GLTPARM3" _
            & " where YYYYPP >= :PARM1 and YYYYPP <= :PARM2"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "VV", 1)
            .Tables("GLTPARM3").Columns.Add("YYYYWW_LY")
            .Tables("GLTPARM3").Columns.Add("MONTH_NO", GetType(System.Int32))
            .Tables("GLTPARM3").Columns.Add("WEEK_NO", GetType(System.Int32))

            Create_TDA(.Tables.Add, "SOTSREG1", "*", 0)
            Fill_Records("SOTSREG1")

            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0)
            Fill_Records("ICTCOLL1")

            '            & " where SPTCWRXC.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            ASCMAIN1.sql = "Select Distinct SPTCOOP1.CUST_CODE, SPTCWRXW.CHECKBOOK" & vbCrLf _
                & ", SPTCOOP1.BOOKING_NAME, SPTCOOP1.EXPENSE_TYPE_CODE, SPTCOOP1.OPS_YYYYWW, SPTCOOP1.VEHICLE_CODE, SPTCOOP1.EVENT_GROUP_NO" & vbCrLf _
                & " from SPTCOOP1,SPTCOOP3, (" & vbCrLf _
                & "Select Distinct ICTCOLL1.COLLECTION_CODE, SPTCWRXC.CHECKBOOK from ICTCOLL1,SPTCWRXC" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
                & "   and (SPTCWRXC.COLLECTION_GENDER = 'U' or ICTCOLL1.COLLECTION_GENDER = SPTCWRXC.COLLECTION_GENDER)) SPTCWRXW" & vbCrLf _
                & " where SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO" & vbCrLf _
                & "   and SPTCOOP1.VEHICLE_CODE <> 'BF'" & vbCrLf _
                & "   and SPTCOOP1.SEASON_CODE = :PARM1" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE in ('A')" & vbCrLf _
                & "   and SPTCWRXW.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
                & "   and SPTCOOP1.EXPENSE_TYPE_CODE IN ('VISUAL','RTLEVENTS')"
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select * from SPTSFOC9 where EVENT_GROUP_NO in (" & vbCrLf _
                & "Select Distinct SPTCOOP1.EVENT_GROUP_NO" & vbCrLf _
                & " from SPTCOOP1" & vbCrLf _
                & " where SPTCOOP1.SEASON_CODE = :PARM1" & vbCrLf _
                & "   and SPTCOOP1.VEHICLE_CODE <> 'BF'" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE in ('A')" & vbCrLf _
                & "   and SPTCOOP1.EXPENSE_TYPE_CODE IN ('VISUAL','RTLEVENTS'))"
            Create_TDA(.Tables.Add, "SPTSFOC9", "**", 0, False, "V", 3)

            ASCMAIN1.sql = "Select * from SPTMXWS5 where ROWNUM < 1"
            SPTMXWS5 = ASCMAIN1.Temp_Table

            'ASCMAIN1.sql = "Select * from " & SPTMXWS5
            'Create_TDA(.Tables.Add, "SPTMXWS5", "**", 0, True, "", 4)
            Create_TDA(.Tables.Add, "SPTMXWS5", "*")

            ASCMAIN1.sql = "Select X.*" & vbCrLf _
                & ", DECODE(C_NEW - C_OLD,1,'Added',-1,'Removed',NULL) C_CHG" & vbCrLf _
                & ", CASE WHEN C_NEW = 0 THEN NULL ELSE CASE WHEN C_NEW <> 0 AND NVL(S_OLD,'?') <> NVL(S_NEW,'?') THEN DECODE(C_OLD,0,'','Now ') || DECODE(NVL(S_NEW,'?'),'I','Inactive','A','Active',NVL(S_NEW,'?') || '?') ELSE NULL END END S_CHG" & vbCrLf _
                & ", CASE WHEN NVL(I_OLD,'?') <> NVL(I_NEW,'?') THEN " & vbCrLf _
                & " CASE WHEN C_NEW = 0 THEN NULL ELSE CASE WHEN NVL(I_NEW,'?') = 'C' THEN 'Closed' ELSE 'Opened' END END" & vbCrLf _
                & " ELSE NULL END I_CHG" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", SUM(C_OLD) C_OLD, MAX(S_OLD) S_OLD, MAX(I_OLD) I_OLD" & vbCrLf _
                & ", SUM(C_NEW) C_NEW, MAX(S_NEW) S_NEW, MAX(I_NEW) I_NEW" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", 0 C_OLD, NULL S_OLD, NULL I_OLD" & vbCrLf _
                & ", 1 C_NEW, CUST_STORE_STATUS S_NEW, CUST_STORE_STATUS_IND I_NEW" & vbCrLf _
                & " from " & SPTMXWS5 & " X" & vbCrLf _
                & " union" & vbCrLf _
                & "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ", 1 C_OLD, CUST_STORE_STATUS S_OLD, CUST_STORE_STATUS_IND I_OLD" & vbCrLf _
                & ", 0 C_NEW, NULL S_OLD, NULL I_OLD" & vbCrLf _
                & " from SPTMXWS5 where SELL_CODE = :PARM1 and SEASON_CODE = :PARM2" & vbCrLf _
                & ") group by CUST_CODE, CUST_STORE_NO" & vbCrLf _
                & ") X where C_OLD <> C_NEW OR NVL(S_OLD,'?') <> NVL(S_NEW,'?') OR  NVL(I_OLD,'?') <> NVL(I_NEW,'?')"
            Create_TDA(.Tables.Add, "SPTMXWSC", "**", 0, False, "VV", 2)

        End With

        Fill_Records("SPTCWRXC")

        grdSPTMXWSC.DataSource = dst.Tables("SPTMXWSC")

        grdSPTMXWSX.DataSource = dst.Tables("SPTMXWSX")
        Create_Summary(grdSPTMXWSX, "SEL")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSPTMXWSX.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        grdSPTMXWSX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        If SELL_CODE_this_user <> "" Then
            grdSPTMXWSX.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        End If

        With grdSPTMXWSX.DisplayLayout.Bands(0)
            .Columns("INIT_DATE").Format = "MM/dd/yy HH:mm"
            .Columns("INIT_DATE").Width = 150
            .Columns("LAST_DATE").Format = "MM/dd/yy HH:mm"
            .Columns("LAST_DATE").Width = 150
        End With


        Dim YYYY As String = Mid(ASCMAIN1.CYP, 1, 4)
        Dim NY As String = Format(Val(YYYY) + 1, "0000")
        Dim LY As String = Format(Val(YYYY) - 1, "0000")

        ASCMAIN1.sql = "Select Min (SEASON_CODE) from SPTMXWS1"
        Dim LY_earliest As String = ASCDATA1.GetDataValue
        If LY_earliest <> "" Then
            LY = Mid(LY_earliest, 1, 4)
        End If

        ASCMAIN1.sql = "Select * from ICTSEAS1 where SEASON_YEAR between '" & LY & "' and '" & NY & "'"
        Dim SEASON_CODEs As New List(Of String)
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "SEASON_YEAR DESC,SEASON_TYPE")
            SEASON_CODEs.Add(row.Item("SEASON_CODE"))
        Next

        Absx1.cbeFor("SEASON_CODE").DataSource = SEASON_CODEs


        Prepare_Work_Tables()

        ' Disable the context menu
        'workbookView1.ContextMenuStrip = Nothing

        ' Replace the context menu
        'WorkbookView1.ContextMenuStrip = myContextMenuStrip

        ' Create and add new item to WorkbookView's context menu
        'Dim newItem As ToolStripItem = WorkbookView1.ContextMenuStrip.Items.Add("Merge Cells")

        ' Add event handler
        'AddHandler newItem.Click, AddressOf MenuItemMergeCells_Click

        Dim MenuItem As ToolStripItem = Nothing
        For i As Integer = WorkbookView1.ContextMenuStrip.Items.Count To 1 Step -1 ' Each MenuItem As ToolStripItem In WorkbookView1.ContextMenuStrip.Items
            MenuItem = WorkbookView1.ContextMenuStrip.Items(i - 1)
            If MenuItem.Text = "&Copy" Or MenuItem.Text = "&Paste" Then
                '  If MenuItem.Text = "Cu&t" Or MenuItem.Text = "&Copy" Or MenuItem.Text = "&Paste" Then
            Else
                MenuItem.Visible = False
                WorkbookView1.ContextMenuStrip.Items.Remove(MenuItem)
            End If
            If MenuItem.Text = "&Paste" Then
                ' AddHandler MenuItem.Click, AddressOf MenuItemPaste_Click
            End If
        Next

        MenuItem = WorkbookView1.ContextMenuStrip.Items.Add("Undo")
        AddHandler MenuItem.Click, AddressOf MenuItemUndo_Click
        MenuItem = WorkbookView1.ContextMenuStrip.Items.Add("Copy to All Stores for Customer")
        AddHandler MenuItem.Click, AddressOf MenuItemCopyNote_Click
    End Sub

    Private Sub MenuItemCopyNote_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim item As ToolStripItem = CType(sender, ToolStripItem)
        If item.Text = "Copy to All Stores for Customer" Then
            WorkbookView1.GetLock()
            Try
                '' Merging is only valid for multi-cell ranges
                'If WorkbookView1.RangeSelection.CellCount >= 2 Then
                '    WorkbookView1.RangeSelection.Merge()
                'End If
                '  SpreadsheetGear.Commands.CommandUndoSupport()

                WorkbookView1.ActiveCommandManager.CreateCommandPaste(WorkbookView1.ActiveCell)
                Dim CP As SpreadsheetGear.Commands.Command = New SpreadsheetGear.Commands.CommandRange.PasteSpecial(range, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)

                '    Dim C As SpreadsheetGear.Commands.Command = WorkbookView1.ActiveCommandManager.CreateCommandPaste(WorkbookView1.ActiveCell, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
                ' range = WorkbookView1.ActiveCell
                range = WorkbookView1.RangeSelection

                If EntryMode = "R" Or range.Cells.ColumnCount <> 1 Or range.Cells.CellCount <> 1 Or range.Cells.RowCount <> 1 Then
                    MsgBox("This Option available on Store Sheets with 1 cell selected")
                Else
                    ' MAKE SURE WE ARE IN A CUSTOMER-STORE SHEET

                    If Not xls_STOREs.Contains(range.Worksheet.Name) Then
                        MsgBox("This Option available on Store Sheets with 1 cell selected")
                    Else

                        Dim NOTE As String = range.Value
                        Dim ADDRESS As String = Replace(range.GetAddress(True, True, SpreadsheetGear.ReferenceStyle.A1, False, Nothing), "$", "")

                        If NOTE = "" Or Not ADDRESS.StartsWith("G") Then
                            MsgBox("This Option available on non-empty Notes cells")
                        Else

                            Dim CUST_CODE As String = Split(range.Worksheet.Name, "-")(0)
                            Dim rows() As DataRow = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE & "'", "HAS_BUDGET DESC,CUST_CODE,CUST_STORE_NO")

                            If MsgBox("OK to Copy note '" & NOTE & "' to All " & CUST_CODE & " Stores?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                                Dim error_message As String = ""
                                For Each row As DataRow In rows
                                    Dim C_VALUE As String = row.Item("CUST_CODE")
                                    Dim S_VALUE As String = row.Item("CUST_STORE_NO")
                                    Dim CS As String = C_VALUE & "-" & S_VALUE

                                    If xls_STOREs.Contains(CS) Then
                                        worksheet = workbook.Worksheets(CS)
                                        Try
                                            worksheet.Cells(ADDRESS).Value = NOTE
                                        Catch ex As Exception
                                            If ex.Message <> "Operation is not valid on locked cells." And error_message = "" Then error_message = ex.Message
                                        End Try

                                    End If
                                Next
                                Dim msg As String = "No Errors - Copy was Successful"
                                If error_message <> "" Then msg = "There were errors during this copy." & vbCrLf & "The message could Not be copied to some stores:" & vbCrLf & vbCrLf & error_message
                                MsgBox(msg, MsgBoxStyle.OkOnly, "Copy Complete")

                            End If
                        End If

                    End If
                End If

            Finally
                WorkbookView1.ReleaseLock()
            End Try
        End If
    End Sub

    Private Sub MenuItemUndo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim item As ToolStripItem = CType(sender, ToolStripItem)
        If item.Text = "Undo" Then
            WorkbookView1.GetLock()
            Try
                '' Merging is only valid for multi-cell ranges
                'If WorkbookView1.RangeSelection.CellCount >= 2 Then
                '    WorkbookView1.RangeSelection.Merge()
                'End If
                '  SpreadsheetGear.Commands.CommandUndoSupport()
                WorkbookView1.ActiveCommandManager.Undo()
            Finally
                WorkbookView1.ReleaseLock()
            End Try
        End If
    End Sub
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("SELL_CODE")

                If SELL_CODE_this_user <> "" And SELL_CODE_this_user <> Absx1.txtFor("SELL_CODE").Text Then
                    EMsg &= vbCrLf & "Invalid AE Code Selected"
                End If
                If REGION_CODE_this_user <> "" Then
                    Dim row As DataRow = LookUp("SOTSELL1", Absx1.txtFor("SELL_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCrLf & "Invalid AE Code Selected"
                    Else
                        If row.Item("REGION_CODE") & "" <> REGION_CODE_this_user Then
                            EMsg &= vbCrLf & "Invalid AE Code Selected"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    Dim row As DataRow = LookUp("SPTMXWS1", New String() {SELL_CODE, SEASON_CODE})
                    If row IsNot Nothing Then
                        EMsg &= vbCr & "Record Already Exists for AE " & SELL_CODE & " in Season " & SEASON_CODE
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SPTMXWS1", "SELL_CODE" & ":" & Absx1.txtFor("SELL_CODE").Text) Then
                        Exit Sub
                    End If
                End If

                'Case "Open"
                '    Dim FILENAME As String = ""
                '    Using openFileDialog1 As New OpenFileDialog
                '        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                '        openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
                '        openFileDialog1.RestoreDirectory = True

                '        If openFileDialog1.ShowDialog() = DialogResult.OK Then
                '            FILENAME = openFileDialog1.FileName
                '        End If
                '    End Using

                '    If FILENAME = "" Then
                '        EMsg &= vbCr & "No Workbook Selected"
                '    Else
                '        WorkbookView1.GetLock()
                '        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

                '        workbook = WorkbookView1.ActiveWorkbook
                '        worksheet = workbook.Worksheets(0)

                '        Absx1.txtFor("SELL_CODE").Text = worksheet.Cells(0, 1).Value
                '        If Absx1.cbeFor("SEASON_CODE").Value <> worksheet.Cells(3, 1).Value Then
                '            EMsg &= vbCrLf & "Season must match " & Absx1.cbeFor("SEASON_CODE").Value & ", and does not (" & worksheet.Cells(3, 1).Value & ")"
                '        End If
                '        WorkbookView1.ReleaseLock()

                '        If EMsg = "" Then
                '            Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
                '            Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                '            Dim row As DataRow = LookUp("SPTMXWS1", New String() {SELL_CODE, SEASON_CODE})
                '            If row Is Nothing Then
                '                EMsg &= vbCr & "No Record on file for AE " & SELL_CODE & " in Season " & SEASON_CODE
                '            End If
                '        End If

                '        If EMsg <> "" Then
                '            Validate_Code("SELL_CODE")
                '            If EMsg = "" Then
                '                If Not ASCMAIN1.Logical_Lock("SPTMXWS1", "SELL_CODE" & ":" & Absx1.txtFor("SELL_CODE").Text) Then
                '                    Exit Sub
                '                End If
                '            End If
                '        End If
                '    End If


            Case "Roll-Up"
                If Absx1.txtFor("SELL_CODE").Text <> "" Then
                    EMsg &= vbCr & "Cannot Roll-Up an individual AE"
                End If

                Dim selected_AEs As Integer = dst.Tables("SPTMXWSX").Select("SEL='1'").Length
                If MsgBox("Do you want to view a Consolidated Freelance Budget Worksheet for " _
                          & IIf(selected_AEs = 0, "All AEs", "the " & CStr(selected_AEs) & " AEs selected") & ", Combined?",
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    'EMsg &= vbCr & "You must first select a Customer"
                    EMsg &= vbCr & "Returning to AE Selection"
                Else
                    blnConsolidated = True
                    SELL_CODES_consolidated.Clear()
                    For Each row As DataRow In dst.Tables("SPTMXWSX").Select("SEL='1'")
                        SELL_CODES_consolidated.Add(row.Item("SELL_CODE"))
                    Next
                End If


            Case "Edit", "View"

                'If eItemKey = "View" And Absx1.txtFor("SELL_CODE").Text = "" Then
                '    If MsgBox("You have not selected an AE, and have clicked View." _
                '              & vbCrLf & vbCrLf & "Do you want to view the Consolidated Stock & Sales Plan for All Customers Combined?", _
                '              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                '        EMsg &= vbCr & "You must first select a Customer"
                '    Else
                '        blnConsolidated = True
                '    End If
                'Else
                Validate_Code("SELL_CODE")

                If SELL_CODE_this_user <> "" And (SELL_CODE_this_user <> Absx1.txtFor("SELL_CODE").Text And (cdr IsNot Nothing AndAlso SELL_CODE_this_user <> cdr.Item("SELL_CODE_MGR") & "")) Then
                    EMsg &= vbCrLf & "Invalid AE Code Selected"
                End If
                If REGION_CODE_this_user <> "" Then
                    Dim row As DataRow = LookUp("SOTSELL1", Absx1.txtFor("SELL_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCrLf & "Invalid AE Code Selected"
                    Else
                        If row.Item("REGION_CODE") & "" <> REGION_CODE_this_user Then
                            EMsg &= vbCrLf & "Invalid AE Code Selected"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim SELL_CODE As String = Absx1.txtFor("SELL_CODE").Text
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    Dim row As DataRow = LookUp("SPTMXWS1", New String() {SELL_CODE, SEASON_CODE})
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record on file for AE " & SELL_CODE & " in Season " & SEASON_CODE
                    End If
                End If
                'End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then

                        If ASCMAIN1.USER_CODES.Contains("FS") And Now.DayOfWeek = DayOfWeek.Friday Then
                        Else
                            If Not ASCMAIN1.Logical_Lock("SPTMXWS1", "SELL_CODE" & ":" & Absx1.txtFor("SELL_CODE").Text) Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Update", "Save"

                If Not isAC And load_from_AC Then
                    If automate_Initialization Then
                        chkApproveAC.Checked = True
                    Else
                        If chkApproveAC.Checked Then
                            If MsgBox("You are an AE Updating an AC workbook, and you have checked the Approve AC Workbook option" _
                              & vbCrLf & vbCrLf & "This will save the data into the FLWB for the AC" _
                              & vbCrLf & vbCrLf & " ----- AND -----" _
                              & vbCrLf & vbCrLf & "Update the data into the FLWB for the Managing AE" _
                              & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        Else
                            If MsgBox("WARNING: Your workbook will NOT be updated." _
                              & vbCrLf & vbCrLf & "You are an AE Updating an AC workbook, but you have NOT checked the Approve AC Workbook option." _
                              & vbCrLf & vbCrLf & "This will save the data into the FLWB for the AC (only)." _
                              & vbCrLf & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If


                WorkbookView1.EndEdit()

            Case "Initialize"

                Dim msg As String = "*** IMPORTANT ***" & vbCrLf & vbCrLf & "Make sure that you have already uploaded the following budgets before you initialize:" & vbCrLf & vbCrLf & "1) Retail Sales Budgets" & vbCrLf & "2) Freelance Budgets" & vbCrLf & "3) Weekly Sales % Plans" & vbCrLf & vbCrLf
                If MsgBox(msg & "OK to Initialize FLWBs for all Active AEs?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Case "Delete FLWBs"

                If MsgBox($"OK to PERMANENTLY Delete all FLWBs for Season {Absx1.cbeFor("SEASON_CODE").Text}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                If MsgBox($"Are you SURE that you want to PERMANENTLY Delete all FLWBs for Season {Absx1.cbeFor("SEASON_CODE").Text}?" & vbCrLf & vbCrLf & "WARNING: This action is PERMANENT.", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Roll-Up"
                EntryMode = "R"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Save"
                SaveMode = True
                Update_Record()
                SaveMode = False

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Save XLSX"
                WorkbookView1.GetLock()
                Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("SPFMXWS1.XLSX_NO") & ".XLSX"
                WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                Show_Document(FILENAME)
                WorkbookView1.ReleaseLock()

            Case "Initialize"
                Initialize_FLWBs()

            Case "Delete FLWBs"

                BeginTrans()

                Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
                Dim SEASON_CODE_JIC As String = ASCMAIN1.Next_Control_No("SPTMXWS1.DELETE")
                For Each TABLE_NAME As String In New String() _
                    {"SPTMXWS1", "SPTMXWS2", "SPTMXWS3", "SPTMXWS4", "SPTMXWS5", "SPTMXAC2", "SPTMXAC3", "SPTMXAC4"}
                    ASCMAIN1.sql = $"Update {TABLE_NAME} Set SEASON_CODE = '{SEASON_CODE_JIC}' where SEASON_CODE = '{SEASON_CODE}'"
                    ASCDATA1.ExecuteSQL()
                Next

                TAC.TACMAIN1.Record_Event("SPTMXWS1", SEASON_CODE, Now, ASCMAIN1.USER_ID, "SSNDEL", $"Season Deleted Archived to {SEASON_CODE_JIC}", SEASON_CODE, Me.Name)

                CommitTrans($"Deletion of all FLWBs for Season {Absx1.cbeFor("SEASON_CODE").Text} is Complete")

                Refresh_Documents()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Open").Settings.Enabled = not_iScreenMode
                    .Items("Save XLSX").Visible = ScreenMode And (EntryMode <> "P")

                    .Items("New").Visible = Not ScreenMode AndAlso (grdSPTMXWSX.Rows.Count = 0)

                    .Items("Roll-Up").Visible = Not ScreenMode And Not isAC And (ASCMAIN1.USER_SECURITY_CODEs.Contains("AA")) And (Not ASCMAIN1.USER_CODES.Contains("FS"))

                    If EntryMode = "V" And ScreenMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    .Items("Save").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L" Or EntryMode = "P")
                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L" Or EntryMode = "P")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L" Or EntryMode = "P")

                    .Items("New").Visible = Not .Items("Update").Visible And Not (blnConsolidated And ScreenMode)
                    .Items("Edit").Visible = Not .Items("Update").Visible And Not (blnConsolidated And ScreenMode)
                    .Items("Open").Visible = False ' Not .Items("Update").Visible And Not blnConsolidated
                    .Items("View").Visible = False ' no support for view - how do we prevent confusing ability to modify workbookview while in view mode?

                    .Items("Done").Visible = ScreenMode And (EntryMode = "V" Or EntryMode = "R")
                    ' .Items("Lock Sales Budget").Visible = Not ScreenMode

                    lblFriday.Appearance.ForeColor = Drawing.Color.Red
                    lblFriday.Visible = False
                    If ASCMAIN1.USER_CODES.Contains("FS") Then
                        lblFriday.Visible = True
                        ' .Items("Lock Sales Budget").Visible = False
                        If Now.DayOfWeek = DayOfWeek.Friday Then
                            .Items("New").Visible = False
                            .Items("Edit").Text = "View-Only-Friday"
                            'lblFriday.Appearance.ForeColor = Drawing.Color.Red
                            .Items("Save").Visible = False
                            .Items("Update").Visible = False
                        End If
                    End If
                End With
                .Groups("Store Status Changes").Visible = ScreenMode And (EntryMode = "E")
                .Groups("Display Options").Visible = ScreenMode And (EntryMode <> "P") And Not blnConsolidated
            End With
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            If Not ScreenMode And ASCMAIN1.USER_SECURITY_CODEs.Contains("FR") Then
            Else
                .Items("Initialize").Visible = False
                .Items("Delete FLWBs").Visible = False
            End If

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        chkApproveAC.Visible = False
        If EntryMode = "E" Then
            If Not isAC And load_from_AC Then
                chkApproveAC.Visible = True
            End If
        End If

        grdSPTMXWSX.Visible = Not ScreenMode
        splSSPL.Visible = ScreenMode And EntryMode <> "P"
        grdSPTMXWSP.Visible = ScreenMode And EntryMode = "P"

        spl.Panel1Collapsed = ScreenMode ' (EntryMode = "P")

        lblACUpdatesPending.Visible = False
        If Not isAC AndAlso (EntryMode = "V" Or EntryMode = "E" Or EntryMode = "N") Then

            ASCMAIN1.sql = "Select SPTMXWS1.SELL_CODE from SPTMXWS1,SOTSELL1" & vbCrLf _
                & " where SPTMXWS1.SEASON_CODE = :PARM1" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE = SPTMXWS1.SELL_CODE" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE_MGR = :PARM2" & vbCrLf _
                & "   and DATA_UPDATED = '1'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", New String() {SEASON_CODE, SELL_CODE})
            Dim ACs As String = ""
            For Each row As DataRow In tbl.Select("", "SELL_CODE")
                ACs &= "." & row.Item("SELL_CODE")
            Next

            If ACs <> "" Then
                lblACUpdatesPending.Text = "AC Updates Pending: " & Mid(ACs, 2)
                lblACUpdatesPending.Visible = True
            End If
        End If

        'If EntryMode = "V" Then
        '    WorkbookView1.GetLock()
        '    For I As Integer = 0 To workbook.Worksheets.Count
        '        range = workbook.Worksheets(I).Cells
        '        range.Locked = True
        '    Next
        '    WorkbookView1.ReleaseLock()
        'End If

        If ScreenMode Then
            If ScreenMode And (EntryMode = "E") Then
                If dst.Tables("SPTMXWSC").Rows.Count <> 0 Then
                    Dim TXT As String = ""
                    For Each row As DataRow In dst.Tables("SPTMXWSC").Select("", "CUST_CODE,CUST_STORE_NO")
                        If TXT <> "" Then TXT &= "<br>" ' vbCrLf
                        TXT &= "Customer " & row.Item("CUST_CODE") & ", Store " & row.Item("CUST_STORE_NO")
                        For Each COL As String In New String() {"C_CHG", "S_CHG", "I_CHG"}
                            If COL <> "C_CHG" And row.Item("C_CHG") & "" = "Added" Then
                                ' do nothing 
                            Else
                                Dim H As String = grdSPTMXWSC.DisplayLayout.Bands(0).Columns(COL).Header.Caption
                                If COL = "C_CHG" Then
                                    H = "AE Assignment"
                                End If

                                If row.Item(COL) & "" <> "" Then TXT &= ", " & H & " " & row.Item(COL)
                            End If
                        Next
                    Next

                    Dim ack As Boolean = False
                    Using F As New ASFMSGBF
                        'Do While Not ack
                        F.Show_Formatted_txt("Store Status Changes", TXT, Me)
                        'F.Show_grd(dst.Tables("SPTMXWSC"), Me, "Store Status Changes")
                        'If F.user_option = -1 Then
                        'Else
                        '    ack = True
                        'End If
                        'Loop
                    End Using
                End If
            End If
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTMXWS1", "SPTMXWS2", "SPTMXWS3", "SPTMXWS4", "SPTMXWS5", "SPTMXWSS", "SPTMXWSC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        SELL_CODE = ""
        SEASON_CODE = ""
        blnConsolidated = False
        chkApproveAC.Checked = False

        Absx1.txtFor("SELL_CODE").Text = ""

        Dim YPS As String = Absx1.cbeFor("SEASON_CODE").Value & ""

        If Absx1.cbeFor("SEASON_CODE").Value & "" = "" Then
            Dim YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)
            Absx1.cbeFor("SEASON_CODE").Value = Mid(YP, 1, 4) & IIf(Mid(YP, 5, 2) < "07", "S", "F")
            If YP = "201807" Then
                Absx1.cbeFor("SEASON_CODE").Value = Mid(YP, 1, 4) & IIf(Mid(YP, 5, 2) < "08", "S", "F")
                ' CM email 06/05/2018
                'I know we just closed the month, but I don’t remember the workbook auto-defaulting to the next season this early.  Can we still keep this as 2018S?  I don’t want any AEs inadvertantly downloading a new workbook before they are ready, or before the new coding is in effect. 
            End If
        Else
            Absx1.cbeFor("SEASON_CODE").Value = ""
            Absx1.cbeFor("SEASON_CODE").Value = YPS
        End If

        If ASCMAIN1.USER_CODES.Contains("FS") And SELL_CODE_this_user <> "" Then
            Absx1.txtFor("SELL_CODE").Text = ""
            Absx1.txtFor("SELL_CODE").Text = SELL_CODE_this_user
            Set_Read_Only_for_ctl(Absx1.txtFor("SELL_CODE"), True)
        Else
            Refresh_Documents()
        End If

        chkAutofitNotes.Checked = False
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        SELL_CODE = HFs("SELL_CODE")

        Dim rowSOTSELL1_THIS_MX As DataRow = LookUp("SOTSELL1", SELL_CODE)
        load_from_AC = False
        If isAC Or (Not blnConsolidated And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V") And rowSOTSELL1_THIS_MX IsNot Nothing AndAlso rowSOTSELL1_THIS_MX.Item("SELL_TYPE") = "AC") Then
            load_from_AC = True
        End If

        'Create_ARTCUST2()
        'Fill_Records("ARTCUST2")

        ASCMAIN1.sql = "SELECT MAX (OPS_YYYYWW) YW_LAST_CWRX FROM SPTCWRX2"
        YW_LAST_CWRX = ASCDATA1.GetDataValue

        SEASON_CODE = HFs("SEASON_CODE")

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        SEASON_TYPE = rowICTSEAS1.Item("SEASON_TYPE")
        SEASON_YEAR = rowICTSEAS1.Item("SEASON_YEAR")

        Dim SEASON_CODE_LP As String = SEASON_YEAR & IIf(SEASON_TYPE = "S", "06", "12")
        useFrozenAlignment = (ASCMAIN1.CYP > SEASON_CODE_LP)
        frozenAlignment_YP = SEASON_CODE_LP

        SEASON_CODE_LY = Format(Val(SEASON_YEAR) - 1, "0000") & SEASON_TYPE

        Fill_Records("SPTMXWS0", SEASON_CODE)
        Fill_Records("SPTCOOPX", SEASON_CODE)
        Fill_Records("SPTSFOC9", SEASON_CODE)

        ReDim YPs(6) ' 1 is either Jan or Jul
        If SEASON_TYPE = "S" Then
            YPs(1) = SEASON_YEAR & "01"
        Else
            YPs(1) = SEASON_YEAR & "07"
        End If
        For I As Integer = 2 To 6
            YPs(I) = ASCMAIN1.Period_Calc(YPs(1), I - 1)
        Next

        ' RELOCATED FROM THE TOP OF THIS PROCEDURE SO THAT WE HAD THE YPS FOR THE SEASON
        Create_ARTCUST2()
        Fill_Records("ARTCUST2")

        AYW = ""
        AYWi = 0
        Number_of_Weeks_in_Season = 0

        Fill_Records("GLTPARM3", New String() {YPs(1), YPs(6)})
        ReDim TYWKs(27)
        ReDim LYWKs(27)
        Dim YW_prior As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1)
        Dim WEEK_NO As Integer = 0
        Dim MONTH_NO As Integer = 0
        Dim YYYYMM As String = ""
        For Each rowGLTPARM3 As DataRow In dst.Tables("GLTPARM3").Select("", "YYYYWW")
            WEEK_NO += 1
            Dim YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
            If rowGLTPARM3.Item("YYYYMM") > YYYYMM Then
                YYYYMM = rowGLTPARM3.Item("YYYYMM")
                MONTH_NO += 1
            End If
            TYWKs(WEEK_NO) = YYYYWW
            LYWKs(WEEK_NO) = ASCMAIN1.Week_Calc(YYYYWW, -52)
            If ASCMAIN1.CLIENT = "INT" Then
                LYWKs(WEEK_NO) = Format(Val(Mid(YYYYWW, 1, 4)) - 1, "0000") & Mid(YYYYWW, 5, 2)
            End If
            rowGLTPARM3.Item("YYYYWW_LY") = LYWKs(WEEK_NO)
            rowGLTPARM3.Item("WEEK_NO") = WEEK_NO
            rowGLTPARM3.Item("MONTH_NO") = MONTH_NO

            If TYWKs(WEEK_NO) <= YW_prior Then
                AYW = TYWKs(WEEK_NO)
                AYWi = WEEK_NO
            End If
        Next
        Number_of_Weeks_in_Season = WEEK_NO


        xls_STs.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTCWRXC"), New String() {"CHECKBOOK_GRP"}).Select("ISNULL(CHECKBOOK_GRP,'') <> ''")
            xls_STs.Add(row.Item("CHECKBOOK_GRP"), New List(Of String))
        Next

        COLS_ST = 1 + xls_STs.Count

        isAE_and_Update_AC = False

        If EntryMode = "P" Then
            Prepare_Summary()
        Else

            Dim between_TY As String = "'" & YPs(1) & "' and '" & YPs(6) & "'"

            ASCMAIN1.sql = "" _
                & "Select Distinct SPTMBUD1.CUST_CODE, SPTMBUD1.CUST_STORE_NO" & vbCrLf _
                & "from SPTMBUD1," & ARTCUST2 & " ARTCUST2," & SPTCWRXW & " SPTCWRXW" & vbCrLf _
                & "where SPTMBUD1.OPS_YYYYPP between " & between_TY & vbCrLf _
                & "  and ARTCUST2.CUST_CODE = SPTMBUD1.CUST_CODE" & vbCrLf _
                & "  and ARTCUST2.CUST_STORE_NO = SPTMBUD1.CUST_STORE_NO" & vbCrLf _
                & "  and SPTCWRXW.COLLECTION_CODE = SPTMBUD1.COLLECTION_CODE" & vbCrLf _
                & "  and SPTMBUD1.BUDGET <> 0" & vbCrLf _
                & " union " & vbCrLf _
                & "Select Distinct SPTMXWS4.CUST_CODE, SPTMXWS4.CUST_STORE_NO" & vbCrLf _
                & "from SPTMXWS4," & ARTCUST2 & " ARTCUST2" & vbCrLf _
                & "where SPTMXWS4.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & "  and ARTCUST2.CUST_CODE = SPTMXWS4.CUST_CODE" & vbCrLf _
                & "  and ARTCUST2.CUST_STORE_NO = SPTMXWS4.CUST_STORE_NO" & vbCrLf _
                & "  and SPTMXWS4.BUDGET_MOD_REV <> 0"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CUST_CODE As String = row.Item("CUST_CODE")
                Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                If rowARTCUST2 IsNot Nothing Then
                    rowARTCUST2.Item("HAS_BUDGET") = "1"
                End If
            Next

            EnforceConstraints(False)

            If EntryMode = "N" Or EntryMode = "R" Or blnConsolidated Then
                rowSPTMXWS1 = dst.Tables("SPTMXWS1").NewRow
                rowSPTMXWS1.Item("SELL_CODE") = SELL_CODE
                rowSPTMXWS1.Item("SEASON_CODE") = SEASON_CODE

                rowSPTMXWS1.Item("INIT_DATE") = DATETIME_STAMP
                rowSPTMXWS1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSPTMXWS1.Item("LAST_DATE") = DATETIME_STAMP
                rowSPTMXWS1.Item("LAST_OPER") = ASCMAIN1.USER_ID

                XLS_NO = ASCMAIN1.Next_Control_No("SPTMXWS1.XLS_NO")
                rowSPTMXWS1.Item("XLS_NO") = XLS_NO

                dst.Tables("SPTMXWS1").Rows.Add(rowSPTMXWS1)
            Else
                rowSPTMXWS1 = Fill_Record("SPTMXWS1", New String() {SELL_CODE, SEASON_CODE})
                XLS_NO = rowSPTMXWS1.Item("XLS_NO")
            End If



            If EntryMode = "R" Then

                ASCMAIN1.sql = "Select SPTMXWS2.SEASON_CODE, ARTCUST2.REGION_CODE CUST_CODE, NVL(ARTCUST2.SELL_CODE,'000') CUST_STORE_NO" & vbCrLf _
                    & ", SPTMXWS2.CHECKBOOK, SPTMXWS2.MONTH_NO, SPTMXWS2.WEEK_NO" & vbCrLf _
                    & ", Sum (TY_SALES) TY_SALES, Sum (TY_HOURS) TY_HOURS, Sum (TY_SPEND) TY_SPEND" & vbCrLf _
                    & " from SPTMXWS2, " & ARTCUST2 & " ARTCUST2 where SPTMXWS2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                    & " and ARTCUST2.CUST_CODE = SPTMXWS2.CUST_CODE AND ARTCUST2.CUST_STORE_NO = SPTMXWS2.CUST_STORE_NO" & vbCrLf _
                    & " group by SPTMXWS2.SEASON_CODE, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
                    & ", SPTMXWS2.CHECKBOOK, SPTMXWS2.MONTH_NO, SPTMXWS2.WEEK_NO"
                Fill_Records("SPTMXWS2", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select SPTMXWS3.SEASON_CODE, ARTCUST2.REGION_CODE CUST_CODE, NVL(ARTCUST2.SELL_CODE,'000') CUST_STORE_NO" & vbCrLf _
                    & ", SPTMXWS3.MONTH_NO, SPTMXWS3.WEEK_NO" & vbCrLf _
                    & ", Max (NOTES) NOTES" & vbCrLf _
                    & " from SPTMXWS3, " & ARTCUST2 & " ARTCUST2 where SPTMXWS3.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                    & " and ARTCUST2.CUST_CODE = SPTMXWS3.CUST_CODE AND ARTCUST2.CUST_STORE_NO = SPTMXWS3.CUST_STORE_NO" & vbCrLf _
                    & " group by SPTMXWS3.SEASON_CODE, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
                    & ", SPTMXWS3.MONTH_NO, SPTMXWS3.WEEK_NO"
                Fill_Records("SPTMXWS3", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select SPTMXWS4.SEASON_CODE, ARTCUST2.REGION_CODE CUST_CODE, NVL(ARTCUST2.SELL_CODE,'000') CUST_STORE_NO" & vbCrLf _
                    & ", SPTMXWS4.CHECKBOOK" & vbCrLf _
                    & ", Sum (BUDGET_RTL) BUDGET_RTL, Sum (BUDGET_MOD) BUDGET_MOD" & vbCrLf _
                    & ", Sum (BUDGET_RTL_REV) BUDGET_RTL_REV, Sum (BUDGET_MOD_REV) BUDGET_MOD_REV" & vbCrLf _
                    & ", Max (BUDGET_RATE) BUDGET_RATE" & vbCrLf _
                    & " from SPTMXWS4, " & ARTCUST2 & " ARTCUST2 where SPTMXWS4.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                    & " and ARTCUST2.CUST_CODE = SPTMXWS4.CUST_CODE AND ARTCUST2.CUST_STORE_NO = SPTMXWS4.CUST_STORE_NO" & vbCrLf _
                    & " group by SPTMXWS4.SEASON_CODE, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
                    & ", SPTMXWS4.CHECKBOOK"
                Fill_Records("SPTMXWS4", "", True, ASCMAIN1.sql)

            Else

                Dim sqlw As String = " where SEASON_CODE = '" & SEASON_CODE & "'" _
                     & " and (CUST_CODE, CUST_STORE_NO) in (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")"
                For Each TABLE_NAME As String In New String() {"SPTMXWS2", "SPTMXWS3", "SPTMXWS4"}
                    ASCMAIN1.sql = "Select * from " & TABLE_NAME & sqlw

                    If load_from_AC Then
                        Dim TAC As String = "SPTMXAC" & Mid(TABLE_NAME, 8, 1)
                        ASCMAIN1.sql = "Select * from " & TAC & sqlw
                    End If

                    Fill_Records(TABLE_NAME, "", True, ASCMAIN1.sql)

                Next
            End If


            If Not blnConsolidated And (EntryMode <> "R") Then
                rowSOTSELL1 = Fill_Record("SOTSELL1", SELL_CODE)
            End If

            EnforceConstraints(True)

            SEASON_SLS_BUD_LOCKED = "0"
            Dim rowSPTMXWSS As DataRow = LookUp("SPTMXWSS", SEASON_CODE)
            If EntryMode = "N" Or EntryMode = "R" Or blnConsolidated Then
                ' NO NEED TO SET THIS FLAG
            Else

                If rowSPTMXWSS IsNot Nothing _
                AndAlso rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") & "" = "1" _
                AndAlso Format(rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED_DATE"), "yyyyMMddHHmmss") <
                        Format(rowSPTMXWS1.Item("LAST_DATE"), "yyyyMMddHHmmss") Then
                    SEASON_SLS_BUD_LOCKED = "1"
                End If

            End If

            If EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V" Or EntryMode = "R" Then
                Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
                If ASCMAIN1.Running_in_VS Then
                    FILENAME = "C:\Share\INT\Templates\" & Me.Name & ".xlsx"
                End If
                WorkbookView1.GetLock()
                WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

                Dim X As SpreadsheetGear.Commands.CommandManager = New MyCommandManager(WorkbookView1.ActiveWorkbookSet)

                XLS_Validation(True)
                XLS_STD()
                XLS_Refresh_Checkbooks()
                XLS_SubTotals()
                Set_Month_Headings(SEASON_CODE)
                XLS_Refresh_Stores()

                WorkbookView1.ReleaseLock()

                Get_Budgets()
                Get_Actuals()
            Else
                WorkbookView1.GetLock()
                XLS_Validation(False)
                ' do we need to do XLS_STD here?
                WorkbookView1.ReleaseLock()
            End If

            If EntryMode = "L" Then
            Else
                WorkbookView1.GetLock()
                worksheet = workbook.Worksheets(0)
                worksheet.Select()
                worksheet.Cells(2, 1).Value = XLS_NO
                worksheet.Cells(2, 2).Value = DATETIME_STAMP
                WorkbookView1.ReleaseLock()
            End If

            If EntryMode = "N" Or EntryMode = "E" Then
                ASCMAIN1.sql = "Truncate Table " & SPTMXWS5
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = $"Select NVL(ARTCUST2.SELL_CODE,'{SELL_CODE}') SELL_CODE, '" & SEASON_CODE & "' SEASON_CODE, ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_STATUS" & vbCrLf _
                    & ", CASE WHEN NVL(HC_OPENED,0) <> 0 THEN 'O' ELSE CASE WHEN NVL(HC_CLOSED,0) <> 0 THEN 'C' ELSE 'N' END END CUST_STORE_STATUS_IND" & vbCrLf _
                    & " from " & ARTCUST2 & " ARTCUST2,(" & vbCrLf _
                    & "Select CUST_CODE, CUST_STORE_NO" & vbCrLf _
                    & ", SUM (CASE WHEN OPS_YYYYPP_CLOSED IS NULL AND OPS_YYYYPP_OPENED IS NOT NULL THEN 1 ELSE 0 END) HC_OPENED" & vbCrLf _
                    & ", SUM (CASE WHEN OPS_YYYYPP_CLOSED IS NOT NULL OR OPS_YYYYPP_OPENED IS NULL THEN 1 ELSE 0 END) HC_CLOSED" & vbCrLf _
                    & " from SATAUTH1 WHERE (CUST_CODE, CUST_STORE_NO) IN (SELECT CUST_CODE, CUST_STORE_NO FROM " & ARTCUST2 & ")" & vbCrLf _
                    & " group by CUST_CODE, CUST_STORE_NO" & vbCrLf _
                    & ") X WHERE ARTCUST2.CUST_CODE = X.CUST_CODE (+) AND ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO (+)"
                ASCDATA1.ExecuteSQL("Insert into " & SPTMXWS5 & " " & ASCMAIN1.sql)

                ASCMAIN1.sql = "Select * from " & SPTMXWS5
                Fill_Records("SPTMXWS5", "", True, ASCMAIN1.sql)

                If EntryMode = "E" Then
                    Fill_Records("SPTMXWSC", New String() {SELL_CODE, SEASON_CODE})
                    Sort_grdColumns(grdSPTMXWSC, "CUST_CODE,CUST_STORE_NO")
                End If
            End If

            WorkbookView1.GetLock()
            For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
                If worksheet.ProtectContents Then worksheet.Unprotect(XLS_PWD)
                worksheet.Protect(XLS_PWD)
            Next
            WorkbookView1.ReleaseLock()

            WorkbookView1.GetLock()
            Try
                For Each ws As SpreadsheetGear.IWorksheet In workbook.Worksheets

                    If ws.ProtectContents Then ws.Unprotect(XLS_PWD)

                    If ws.Name = "Summary" Then
                        ws.Protect(
                password:=XLS_PWD,
                protectDrawingObjects:=True,
                protectScenarios:=True,
                userInterfaceOnly:=True,
                allowFormattingCells:=False,
                allowFormattingColumns:=True,
                allowFormattingRows:=True,
                allowInsertingColumns:=False,
                allowInsertingRows:=False,
                allowInsertingHyperlinks:=False,
                allowDeletingColumns:=False,
                allowDeletingRows:=False,
                allowSorting:=False,
                allowFiltering:=False,
                allowUsingPivotTables:=False
            )
                    Else
                        ws.Protect(
                password:=XLS_PWD,
                protectDrawingObjects:=True,
                protectScenarios:=True,
                userInterfaceOnly:=True
            )
                    End If

                Next
            Finally
                WorkbookView1.ReleaseLock()
            End Try


            If LOCK_SHEETS.Count > 0 Then
                WorkbookView1.GetLock()

                If EntryMode = "E" AndAlso Not isAC Then isAE_and_Update_AC = True

                Dim ACs As New Dictionary(Of String, Boolean)
                For Each CS As String In LOCK_SHEETS.Keys
                    'Dim C_VALUE As String = Split(CS, "-")(0)
                    'Dim S_VALUE As String = Split(CS, "-")(1)
                    Dim AC As String = LOCK_SHEETS(CS)
                    If Not ACs.ContainsKey(AC) Then
                        Dim row As DataRow = LookUp("SPTMXWS1", New String() {AC, SEASON_CODE})
                        If row IsNot Nothing AndAlso row.Item("DATA_UPDATED") & "" = "1" Then
                            ACs.Add(AC, False)
                        Else
                            ACs.Add(AC, ASCMAIN1.Logical_Lock("SPTMXWS1", "SELL_CODE" & ":" & AC, False, False, False))
                        End If
                    End If
                    Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(CS)
                    If ws.ProtectContents Then ws.Unprotect(XLS_PWD)
                    ws.Cells("D1").Value = "Territory Recap (AC " & LOCK_SHEETS(CS) & ")"
                    ws.Cells("D1").Font.Color = SpreadsheetGear.Colors.Red

                    If ACs(AC) Then
                        ws.Tab.Color = SpreadsheetGear.Colors.Yellow
                        ws.Protect(XLS_PWD)
                    Else
                        isAE_and_Update_AC = False
                        ws.Cells.Locked = True
                        ws.Protect(XLS_PWD)
                        ws.Tab.Color = SpreadsheetGear.Colors.Red
                    End If

                Next
                WorkbookView1.ReleaseLock()
            End If


        End If

        ASCMAIN1.Progress("")
    End Sub


    Sub Update_Record()

        For Each TABLE_NAME As String In New String() {"SPTMXWS2", "SPTMXWS3", "SPTMXWS4"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        WorkbookView1.GetLock()

        For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "CUST_CODE, CUST_STORE_NO")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim CS As String = CUST_CODE & "-" & CUST_STORE_NO

            worksheet = workbook.Worksheets(CS)

            For CBi2 As Integer = 1 To CBi

                Dim BUDGET_RTL As Decimal = Val(worksheet.Cells(3, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDRTL).Value & "")
                Dim BUDGET_MOD As Decimal = Val(worksheet.Cells(3, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDMOD).Value & "")
                Dim BUDGET_RTL_REV As Decimal = Val(worksheet.Cells(4, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDRTL).Value & "")
                Dim BUDGET_MOD_REV As Decimal = Val(worksheet.Cells(4, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDMOD).Value & "")
                Dim BUDGET_RATE As Decimal = Val(worksheet.Cells(4, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colLYRACT).Value & "")
                Dim BUDGET_RATE_FORMULA As String = worksheet.Cells(4, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colLYRACT).Formula & ""

                If BUDGET_RTL <> 0 Or BUDGET_MOD <> 0 Or BUDGET_RTL_REV <> 0 Or BUDGET_MOD_REV <> 0 Or Not BUDGET_RATE_FORMULA.StartsWith("=") Then
                    Dim rowSPTMXWS4 As DataRow = dst.Tables("SPTMXWS4").NewRow
                    With rowSPTMXWS4
                        .Item("SEASON_CODE") = SEASON_CODE
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                        .Item("CHECKBOOK") = xls_CBs(CBi2 - 1)
                        .Item("BUDGET_RTL") = BUDGET_RTL
                        .Item("BUDGET_MOD") = BUDGET_MOD
                        .Item("BUDGET_RTL_REV") = BUDGET_RTL_REV
                        .Item("BUDGET_MOD_REV") = BUDGET_MOD_REV
                        .Item("BUDGET_RATE") = BUDGET_RATE
                    End With
                    dst.Tables("SPTMXWS4").Rows.Add(rowSPTMXWS4)
                End If
            Next

            For MONTH_NO As Integer = 1 To 6
                Dim NOTES As String = worksheet.Cells(SROW_Months + MONTH_NO, 6).Value & ""
                If NOTES <> "" Then
                    If NOTES.Length > 100 Then NOTES = Mid(NOTES, 1, 100)
                    Dim rowSPTMXWS3 As DataRow = dst.Tables("SPTMXWS3").NewRow
                    With rowSPTMXWS3
                        .Item("SEASON_CODE") = SEASON_CODE
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                        .Item("MONTH_NO") = MONTH_NO
                        .Item("WEEK_NO") = 0
                        .Item("NOTES") = NOTES
                    End With
                    dst.Tables("SPTMXWS3").Rows.Add(rowSPTMXWS3)
                End If
            Next

            For WEEK_NO As Integer = 1 To Number_of_Weeks_in_Season
                Dim YW As String = TYWKs(WEEK_NO)
                Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YW)
                Dim MONTH_NO As Integer = Val(rowGLTPARM3.Item("MONTH_NO") & "")

                For CBi2 As Integer = 1 To CBi

                    Dim TY_SALES As Decimal = Val(worksheet.Cells(SROW_Weeks + WEEK_NO, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDRTL).Value & "")
                    Dim TY_HOURS As Decimal = Val(worksheet.Cells(SROW_Weeks + WEEK_NO, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDHRS).Value & "")
                    Dim TY_SPEND As Decimal = Val(worksheet.Cells(SROW_Weeks + WEEK_NO, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDMOD).Value & "")

                    If TY_SALES <> 0 Or TY_HOURS <> 0 Or TY_SPEND <> 0 Then
                        Dim rowSPTMXWS2 As DataRow = dst.Tables("SPTMXWS2").NewRow
                        With rowSPTMXWS2
                            .Item("SEASON_CODE") = SEASON_CODE
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("CUST_STORE_NO") = CUST_STORE_NO
                            .Item("CHECKBOOK") = xls_CBs(CBi2 - 1)
                            .Item("MONTH_NO") = MONTH_NO
                            .Item("WEEK_NO") = WEEK_NO
                            .Item("TY_SALES") = TY_SALES
                            .Item("TY_HOURS") = TY_HOURS
                            .Item("TY_SPEND") = TY_SPEND
                        End With
                        dst.Tables("SPTMXWS2").Rows.Add(rowSPTMXWS2)
                    End If

                Next

                Dim NOTES As String = worksheet.Cells(SROW_Weeks + WEEK_NO, 6).Value & ""
                If NOTES <> "" Then
                    If NOTES.Length > 100 Then NOTES = Mid(NOTES, 1, 100)
                    Dim rowSPTMXWS3 As DataRow = dst.Tables("SPTMXWS3").NewRow
                    With rowSPTMXWS3
                        .Item("SEASON_CODE") = SEASON_CODE
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                        .Item("MONTH_NO") = MONTH_NO
                        .Item("WEEK_NO") = WEEK_NO
                        .Item("NOTES") = NOTES
                    End With
                    dst.Tables("SPTMXWS3").Rows.Add(rowSPTMXWS3)
                End If
            Next
        Next

        WorkbookView1.ReleaseLock()

        BeginTrans()

        ' Dim rowSPTMXWS1 As DataRow = dst.Tables("SPTMXWS1").Rows.Find(New String() {SELL_CODE, SEASON_CODE})
        If isAC Then
            rowSPTMXWS1.Item("DATA_UPDATED") = "1"
        Else
            If load_from_AC And chkApproveAC.Checked Then
                rowSPTMXWS1.Item("DATA_UPDATED") = "0"
            End If
        End If

        Dim sqld As String = "SELL_CODE = '" & SELL_CODE & "' and SEASON_CODE = '" & SEASON_CODE & "'"
        INIT_LAST("SPTMXWS1")

        Update_Record_TDA("SPTMXWS1", sqld)
        Update_Record_TDA("SPTMXWS5", sqld)

        sqld = "SEASON_CODE = '" & SEASON_CODE & "'" _
            & " and (CUST_CODE, CUST_STORE_NO) in " _
            & " (Select CUST_CODE, CUST_STORE_NO from " & ARTCUST2 & ")"


        If isAC Or load_from_AC Or isAE_and_Update_AC Then
            ' if calling up an AC when you are not an AC,
            ' then we really need to update the ACs tables as well as the real ones

            For Each T As String In New String() {"SPTMXWS2", "SPTMXWS3", "SPTMXWS4"}
                Dim TAC As String = "SPTMXAC" & Mid(T, 8, 1)
                dst.Tables(TAC).Rows.Clear()
                For Each ROW As DataRow In dst.Tables(T).Select("")
                    dst.Tables(TAC).Rows.Add(ROW.ItemArray)
                Next
            Next

            Update_Record_TDA("SPTMXAC2", sqld)
            Update_Record_TDA("SPTMXAC3", sqld)
            Update_Record_TDA("SPTMXAC4", sqld)

        End If

        If Not isAC And (Not load_from_AC Or chkApproveAC.Checked) Then



            Update_Record_TDA("SPTMXWS2", sqld)
            Update_Record_TDA("SPTMXWS3", sqld)
            Update_Record_TDA("SPTMXWS4", sqld)
        End If


        Dim z As String = "Update Complete"
        If SaveMode Then
            z = "Save Complete"
        End If

        If automate_Initialization Then
            z = ""
        End If

        CommitTrans(z)
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
        Load_Popup_Menu(grdSPTMXWSX, "SS", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Select All for Region")

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
                Case "grdSPTMXWSX"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                    tlb_btn = DirectCast(tlb_pop.Tools("Select All for Region"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim REGION_CODE As String = grd.ActiveRow.Cells("REGION_CODE").Value & ""
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Select All for Region " & REGION_CODE
                    End If

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTMXWSX.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All for Region"
                Dim REGION_CODE As String = grd.ActiveRow.Cells("REGION_CODE").Value & ""
                Dim sqlw As String = "REGION_CODE = '" & REGION_CODE & "'"
                For Each row As DataRow In dst.Tables("SPTMXWSX").Select(sqlw)
                    row.Item("SEL") = "1"
                Next

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
            Case "SELL_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

            'Case "HC_CODE"
            '    If Absx1.txtFor("HC_CODE").Text <> "" Then
            '        If LookUp("ICTCOLL0", Absx1.txtFor("HC_CODE").Text) IsNot Nothing Then
            '            XLS_Refresh_HC(Absx1.txtFor("HC_CODE").Text)
            '        End If
            '    End If
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SELL_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("SELL_CODE").Text <> "" Then
                        LookUp("SOTSELL1", Absx1.txtFor("SELL_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub cbe_ValueChanged(sender As Object, e As EventArgs)
        MyBase.cbe_ValueChanged(sender, e)
        Refresh_Documents()
    End Sub

#End Region

    Overrides Sub Prepare_for_View_Lookup_Special(
 ByVal ctl As Control,
 ByVal COLUMN_NAME As String,
 Optional ByRef sql_where As String = "",
 Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "SELL_CODE"
                If SELL_CODE_this_user <> "" Then
                    sql_where = "SELL_CODE = '" & SELL_CODE_this_user & "'"
                ElseIf REGION_CODE_this_user <> "" Then
                    sql_where = "REGION_CODE = '" & REGION_CODE_this_user & "'"
                End If
        End Select
    End Sub

    Sub Set_Month_Headings(SEASON_CODE As String)

        Dim worksheetRecap As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")

        Dim M As New Dictionary(Of String, Integer)

        For i As Integer = 1 To 6
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YPs(i))
            Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")

            worksheetRecap.Cells(SROW_Months + i, 3).Value = Mid(LEGEND, 10, 3)
            M.Add(YPs(i), i)
        Next

        For i As Integer = 1 To 27
            Dim YW As String = TYWKs(i)
            If YW <> "" Then
                Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(YW)
                Dim WEEK_END_DATE As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                Dim WEEK_BEG_DATE As Date = WEEK_END_DATE.AddDays(-6)
                Dim MONTH_NO As Integer = Val(rowGLTPARM3.Item("MONTH_NO") & "")
                Dim rowSPTMXWS0 As DataRow = dst.Tables("SPTMXWS0").Rows.Find(New Object() {SEASON_CODE, Format(i, "00")})

                Dim PCT_M As Decimal = 0
                Dim PCT_W As Decimal = 0

                If rowSPTMXWS0 IsNot Nothing Then
                    PCT_M = Val(rowSPTMXWS0.Item("PCT_M") & "") / 100
                    PCT_W = Val(rowSPTMXWS0.Item("PCT_W") & "") / 100
                End If
                worksheetRecap.Cells(SROW_Weeks + i, 1).Value = "M" & M(rowGLTPARM3.Item("YYYYMM"))
                worksheetRecap.Cells(SROW_Weeks + i, 3).Value = Format(WEEK_BEG_DATE, "MM/dd")
                worksheetRecap.Cells(SROW_Weeks + i, 4).Value = PCT_M ' Mens Penetration %
                worksheetRecap.Cells(SROW_Weeks + i, 5).Value = PCT_W ' Womens Penetration %

                ' section below is unnecessary since each workbook handles these columns with locking and formulae, and the section below is missing -1 adj for SCOL_Recap
                'If i <= AYWi Then
                '    For Each c As Integer In New Integer() {colBUDRTL, colBUDHRS}
                '        worksheetRecap.Cells(SROW_Weeks + i, SCOL_Recap + c).Locked = True
                '        worksheetRecap.Cells(SROW_Weeks + i, SCOL_Recap + c).Interior.Color = SpreadsheetGear.Colors.White
                '    Next
                'End If

            Else
                If i = 27 Then
                    worksheetRecap.Cells(SROW_Weeks + i, 0).EntireRow.Delete()
                End If
            End If
        Next

        worksheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + 1 + 1).Activate()
        worksheetRecap.WindowInfo.FreezePanes = True
    End Sub

    Sub Refresh_Documents()
        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
        Fill_Records("SPTMXWSX", SEASON_CODE)
        Sort_grdColumns(grdSPTMXWSX, "SELL_CODE")
        grdSPTMXWSX.Text = "Model Expense Worksheets for " & SEASON_CODE

        If ASCMAIN1.USER_CODES.Contains("FS") Then
            btnLockSalesBudget.Visible = False
        Else
            Dim rowSPTMXWSS As DataRow = LookUp("SPTMXWSS", SEASON_CODE)
            If rowSPTMXWSS IsNot Nothing AndAlso rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") & "" = "1" Then
                Dim SEASON_SLS_BUD_LOCKED_DATE As Date = rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED_DATE")
                UltraLabel3.Text = "Season - Sales Budget Locked " & Format(SEASON_SLS_BUD_LOCKED_DATE, "MM/dd/yy HH:mm")
                UltraLabel3.Appearance.ForeColor = System.Drawing.Color.Empty
                btnLockSalesBudget.Visible = False
                btnUnLockSalesBudget.Visible = (Format(SEASON_SLS_BUD_LOCKED_DATE.AddDays(30), "yyyyMMdd") > Format(Now, "yyyyMMdd"))
            Else
                UltraLabel3.Text = "Season - Sales Budget NOT Locked"
                UltraLabel3.Appearance.ForeColor = System.Drawing.Color.Red
                btnLockSalesBudget.Visible = True
                btnUnLockSalesBudget.Visible = False
            End If
        End If

        If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("FR") Then
            btnLockSalesBudget.Visible = False
            btnUnLockSalesBudget.Visible = False
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("Initialize").Visible = (grdSPTMXWSX.Rows.Count = 0 And Not ScreenMode And ASCMAIN1.USER_SECURITY_CODEs.Contains("FR"))
            .Items("Delete FLWBs").Visible = (grdSPTMXWSX.Rows.Count > 0 And Not ScreenMode And ASCMAIN1.USER_SECURITY_CODEs.Contains("FR"))
        End With

    End Sub

    Sub XLS_Validation(isTemplate As Boolean)
        Dim sheet_valid As Boolean = True
        Dim sheet_error_msg As String = ""
        XLSR.Clear()
        ' Delete_Sheets.Clear()
        xls_STOREs.Clear()

        workbook = WorkbookView1.ActiveWorkbook

        If workbook Is Nothing OrElse workbook.Worksheets.Count < 2 OrElse workbook.Worksheets(0).Name <> "Summary" OrElse workbook.Worksheets(1).Name <> "Recap" Then
            sheet_error_msg = "Workbook does not contain at least 2 sheets beginning with the Summary Sheet and the Recap Sheet"
        Else
            'If isTemplate Then
            '    Delete_Sheets.Add(workbook.Worksheets(1).Name)
            '    Delete_Sheets.Add(workbook.Worksheets(2).Name)
            'End If

            For i As Integer = 0 To workbook.Worksheets.Count - 1
                Dim sheet_name As String = workbook.Worksheets(i).Name
                If sheet_name <> "Summary" And sheet_name <> "Recap" Then
                    ' And Not Delete_Sheets.Contains(sheet_name) Then
                    xls_STOREs.Add(sheet_name)
                End If
                If sheet_name = "Recap" Then
                    ReDim CVG_PCT(27)
                    For iCVG As Integer = 1 To 27
                        Dim c As String = "O" & CStr(19 + iCVG - 1)
                        Dim CVG_PCT_FORMULA As String = workbook.Worksheets(sheet_name).Cells(c).Formula
                        '=IF(K19=0,0,+M19*100/K19)
                        If InStr(CVG_PCT_FORMULA, "*") <> 0 And InStr(CVG_PCT_FORMULA, "/") <> 0 _
                            And InStr(CVG_PCT_FORMULA, "*") < InStr(CVG_PCT_FORMULA, "/") Then

                            CVG_PCT_FORMULA = Split(CVG_PCT_FORMULA, "*")(1)
                            CVG_PCT_FORMULA = Split(CVG_PCT_FORMULA, "/")(0)
                            CVG_PCT(iCVG) = CStr(Val(CVG_PCT_FORMULA))
                        Else
                            CVG_PCT(iCVG) = "100"
                        End If
                    Next
                End If
            Next
            worksheet = workbook.Worksheets("Summary")
        End If

        If sheet_error_msg = "" Then
        End If
        sheet_valid = (sheet_error_msg = "")
    End Sub

    Sub XLS_SubTotals()

        If xls_STs.Count = 0 Then Exit Sub

        Dim workSheetRecap As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")

        Dim STi As Integer = 0
        For Each ST As String In xls_STs.Keys

            STi += 1
            Dim c As Integer = SCOL_Recap + COLS_Recap + STi

            Dim CBs As String = ""
            Dim XLA As String = ""
            For Each row As DataRow In dst.Tables("SPTCWRXC").Select("CHECKBOOK_GRP = '" & ST & "'", "CHECKBOOK")
                Dim CHECKBOOK As String = row.Item("CHECKBOOK")
                Dim CBi2 As Integer = Val(row.Item("CBI") & "")
                CBs &= "+" & CHECKBOOK
                XLA &= "+" & Excel_Cell0(SROW_Months + 1, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colBUDHRS)
            Next

            workSheetRecap.Cells(0, c).Value = "Sub-" & ST
            workSheetRecap.Cells(1, c).Value = Mid(CBs, 2)
            workSheetRecap.Cells(SROW_Months, c).Value = "Hrs"
            workSheetRecap.Cells(SROW_Months + 1, c).Formula = "=" & Mid(XLA, 2)
            workSheetRecap.Cells(SROW_Weeks, c).Value = "Hrs"

            ' Copy Cell for Month 1 sub-total ...
            rangeCopyFrom = workSheetRecap.Range(SROW_Months + 1, c)

            ' ... to months 2 thru 6
            rangePaste_To = workSheetRecap.Range(SROW_Months + 1, c, SROW_Months + 6, c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

            ' ... to all weeks in the season
            rangePaste_To = workSheetRecap.Range(SROW_Weeks + 1, c, SROW_Weeks + Number_of_Weeks_in_Season, c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White


            ' Copy the Formats for Hours (in the monthly section) to the sub-totals columns
            rangeCopyFrom = workSheetRecap.Range(SROW_Months + 0, SCOL_Recap + colBUDHRS,
                                                 SROW_Months + 6 + 3, SCOL_Recap + colBUDHRS)
            rangePaste_To = workSheetRecap.Range(SROW_Months + 0, c,
                                                 SROW_Months + 6 + 3, c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formats, SpreadsheetGear.PasteOperation.None, False, False)

            ' Copy the Formats for Hours (in the weekly section) to the sub-totals columns
            rangeCopyFrom = workSheetRecap.Range(SROW_Weeks + 0, SCOL_Recap + colBUDHRS,
                                                 SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + colBUDHRS)
            rangePaste_To = workSheetRecap.Range(SROW_Weeks + 0, c,
                                                 SROW_Weeks + Number_of_Weeks_in_Season, c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formats, SpreadsheetGear.PasteOperation.None, False, False)


            ' Copy the Totals for Hours (in the monthly section) to the sub-totals columns
            For Each irow As Integer In New Integer() {1, 3}
                rangeCopyFrom = workSheetRecap.Range(SROW_Months + 6 + irow, SCOL_Recap + colBUDHRS,
                                                     SROW_Months + 6 + irow, SCOL_Recap + colBUDHRS)
                rangePaste_To = workSheetRecap.Range(SROW_Months + 6 + irow, c,
                                                     SROW_Months + 6 + irow, c)
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                rangePaste_To.Locked = True
                rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
            Next
        Next

        range = workSheetRecap.Cells(0, SCOL_Recap + COLS_Recap + 1,
                                     0, SCOL_Recap + COLS_Recap + xls_STs.Count)
        range.Interior.Color = SpreadsheetGear.Colors.LightGray

        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
    End Sub

    Sub XLS_STD()

        If AYWi = 0 Then Exit Sub

        Dim workSheetRecap As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")

        Dim XLA As String = Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + colLYRHRS) & ":" &
                            Excel_Cell0(SROW_Weeks + AYWi, SCOL_Recap + colLYRHRS)
        workSheetRecap.Cells(SROW_Weeks - 1, SCOL_Recap + colLYRHRS).Formula = "=Sum(" & XLA & ")"

        rangeCopyFrom = workSheetRecap.Range(SROW_Weeks - 1, SCOL_Recap + colLYRHRS,
                                             SROW_Weeks - 1, SCOL_Recap + colLYRHRS)
        For Each c As Integer In New Integer() {colLYRACT, colLYRSLS, colBUDRTL, colBUDHRS, colBUDMOD}
            rangePaste_To = workSheetRecap.Range(SROW_Weeks - 1, SCOL_Recap + c,
                                                 SROW_Weeks - 1, SCOL_Recap + c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
        Next

        For Each c As Integer In New Integer() {colPCTRTL, colCVGMOD}
            rangeCopyFrom = workSheetRecap.Range(SROW_Weeks - 3, SCOL_Recap + c,
                                                 SROW_Weeks - 3, SCOL_Recap + c)
            rangePaste_To = workSheetRecap.Range(SROW_Weeks - 1, SCOL_Recap + c,
                                                 SROW_Weeks - 1, SCOL_Recap + c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
        Next
    End Sub

    Sub XLS_Refresh_Checkbooks()

        Dim workSheetRecap As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")
        Dim workSheetSummary As SpreadsheetGear.IWorksheet = workbook.Worksheets("Summary")

        With workSheetSummary
            .Range(0, 0, 0, 3).EntireColumn.Hidden = True
        End With

        With workSheetRecap
            .Range(0, 0, 0, 2).EntireColumn.Hidden = True
        End With

        If (EntryMode = "R") Then
            workSheetSummary.Cells(0, 4).Value = "AE Plans vs Freelance Budget Summary"
            workSheetSummary.Cells(1, 4).Value = "*"
        Else
            workSheetSummary.Cells(0, 4).Value = rowSOTSELL1.Item("SELL_NAME")
            workSheetSummary.Cells(1, 4).Value = SELL_CODE
        End If
        workSheetSummary.Cells(0, 5).Value = SEASON_CODE

        Dim WKZ As String = ""
        If AYWi = 0 Then

        Else
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", AYW)
            WKZ = rowGLTPARM3.Item("LEGEND")
            WKZ = Mid(WKZ, 10, 7)
            WKZ = Mid(WKZ, 1, 3) & " Wk " & Mid(WKZ, 5, 1)

        End If

        Dim WKM As String = ""
        Dim MYW As String = ASCDATA1.GetDataValue("Select MAX (OPS_YYYYWW) from SPTCWRX2")
        If MYW <> "" Then
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", MYW)
            WKM = rowGLTPARM3.Item("LEGEND")
            WKM = Mid(WKM, 10, 7)
            WKM = Mid(WKM, 1, 3) & " Wk " & Mid(WKM, 5, 1)
        End If

        workSheetSummary.Cells(2, 4).Value = SEASON_CODE & IIf(AYWi = 0, "", " Sls Act " & WKZ & "; FL " & WKM) ' CStr(AYWi)

        'workSheetSummary.Cells(3, 4).Value = "Freelance Actualized thru " & WKM
        'workSheetSummary.Cells(3, 4).Font.Color = SpreadsheetGear.Colors.Red

        workSheetSummary.Cells(3, 1).Value = ""
        workSheetSummary.Cells(4, 1).Value = ""


        CBi = 0

        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST * 0 + (CBi) * COLS_Recap + 1).Value = "Total"
        workSheetRecap.Cells(1, SCOL_Recap + COLS_ST * 0 + (CBi) * COLS_Recap + 2).Value = "*"

        workSheetSummary.Cells(0, SCOL_Summary + (CBi) * COLS_Summary + 1 - 1).Value = "Total"
        workSheetSummary.Cells(1, SCOL_Summary + (CBi) * COLS_Summary + 1 - 1).Value = "*"

        Dim XLA As String = ""
        xls_CBs.Clear()


        For Each row As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
            CBi += 1
            row.Item("CBI") = CBi
            Dim CHECKBOOK As String = row.Item("CHECKBOOK")
            Dim CHECKBOOK_DESC As String = row.Item("CHECKBOOK_DESC")

            Dim CHECKBOOK_GRP As String = row.Item("CHECKBOOK_GRP") & ""
            If CHECKBOOK_GRP = "" Then
                ' ADD TO UNGROUPED CHECKBOOKS
                xls_NOT_STs.Add(CHECKBOOK)
            Else
                xls_STs(CHECKBOOK_GRP).Add(CHECKBOOK)
            End If



            Dim rr As String = Excel_Cell0(-1, SCOL_Recap + 1) & ":" &
                               Excel_Cell0(-1, SCOL_Recap + COLS_Recap)
            'rangeCopyFrom = workSheetRecap.Range("G:O")
            rangeCopyFrom = workSheetRecap.Range(rr)
            ' Stop ' needs fixing
            rangePaste_To = workSheetRecap.Range(Excel_Cell0(-1, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + 1) & ":" &
                                                 Excel_Cell0(-1, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + COLS_Recap))
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

            workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + 1).Value = CHECKBOOK_DESC
            workSheetRecap.Cells(1, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + 2).Value = CHECKBOOK
            workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + 0).ColumnWidth = 2


            rangeCopyFrom = workSheetSummary.Range("G:J")
            rangePaste_To = workSheetSummary.Range(Excel_Cell(0, SCOL_Summary + (CBi) * COLS_Summary + 1) & ":" &
                                                   Excel_Cell(0, SCOL_Summary + (CBi) * COLS_Summary + COLS_Summary))
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

            workSheetSummary.Cells(0, SCOL_Summary + (CBi) * COLS_Summary + 1 - 1).Value = CHECKBOOK_DESC
            workSheetSummary.Cells(1, SCOL_Summary + (CBi) * COLS_Summary + 1 - 1).Value = CHECKBOOK
            workSheetSummary.Cells(0, SCOL_Summary + (CBi) * COLS_Summary + 1 - 2).ColumnWidth = 2

            xls_CBs.Add(CHECKBOOK)
            ' XLA = "+S19+AC19+AM19+AW19+BG19+BQ19"
            XLA &= "+" & Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colLYRHRS)

            range = workSheetRecap.Cells(SROW_Months + 0, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES,
                                         SROW_Months + 6, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES)
            range.Clear()
            range.Locked = True
            range.Interior.Color = SpreadsheetGear.Colors.White

            range = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES,
                                         SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES)
            range.Locked = True
            range.Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
            workSheetRecap.Cells(SROW_Weeks + 0, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES).Value = "Events"
            workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPNOTES).EntireColumn.ColumnWidth = 12

        Next

        ' Set the Formula for Total Hours LY as the sum of each checkbook Total Hours LY
        ' this formula will be copied to several ranges
        ' for each range we paste to, we will lock down the range for exiting
        workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + colLYRHRS).Formula = "=" & XLA
        rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + colLYRHRS)

        '  Stop ' the next few lines had + COLS_ST removed from the paste to
        ' Paste the Formula to the Sales Goal & Spread
        rangePaste_To = workSheetRecap.Range(SROW_Months - 4, SCOL_Recap + colBUDRTL,
                                             SROW_Months - 3, SCOL_Recap + colBUDRTL)
        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
        rangePaste_To.Locked = True
        rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

        ' Paste the Formula to the Spend Budget & Spread
        rangePaste_To = workSheetRecap.Range(SROW_Months - 4, SCOL_Recap + colBUDMOD,
                                             SROW_Months - 3, SCOL_Recap + colBUDMOD)
        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
        rangePaste_To.Locked = True
        rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

        ' Paste the Formula to Total Hours - already locked so no need to re-lock
        rangePaste_To = workSheetRecap.Range(SROW_Months - 3, SCOL_Recap + colCVGMOD)
        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
        'rangePaste_To.Locked = True
        'rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

        ' No need to paste to Plan Rate - just lock it down, and re-calculate based on totals
        rangePaste_To = workSheetRecap.Range(SROW_Months - 3, SCOL_Recap + colPLANRATE)
        rangePaste_To.Locked = True
        rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
        'rangePaste_To.Formula = "=IFERROR(" & Excel_Cell0(4, SCOL_Recap + colBUDHRS) & "/" & _
        '                                      Excel_Cell0(4, SCOL_Recap + colBUDMOD) & ",0)"

        ' Paste the Formula to all of the weeks in the Weekly Section
        rangePaste_To = workSheetRecap.Range(SROW_Weeks + 2, SCOL_Recap + colLYRHRS,
                                             SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + colLYRHRS)
        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
        rangePaste_To.Locked = True
        rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

        ' Paste the Formula to Spend LY, Sales LY, Sales TY, Hours TY, and Spend TY
        rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + colLYRHRS,
                                             SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + colLYRHRS)
        For Each c As Integer In New Integer() {colLYRACT, colLYRSLS, colBUDRTL, colBUDHRS, colBUDMOD}
            rangePaste_To = workSheetRecap.Range(SROW_Weeks + 1, SCOL_Recap + c,
                                                 SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
        Next


        If xls_NOT_STs.Count > 0 Then
            For Each CHECKBOOK As String In xls_NOT_STs
                Dim rowSPTCWRXC As DataRow = dst.Tables("SPTCWRXC").Rows.Find(CHECKBOOK)
                Dim CBI As Integer = Val(rowSPTCWRXC.Item("CBI") & "")
                workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + 10).EntireColumn.Hidden = True
                workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_SLS).EntireColumn.Hidden = True
                workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_HRS).EntireColumn.Hidden = True
                workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_CVG).EntireColumn.Hidden = True
            Next
        End If

        ' Hide the ST columns following the Total block - prob should use 

        workSheetRecap.Cells(0, SCOL_Recap + 10).EntireColumn.Hidden = True
        workSheetRecap.Cells(0, SCOL_Recap + colST_SLS).EntireColumn.Hidden = True
        workSheetRecap.Cells(0, SCOL_Recap + colST_HRS).EntireColumn.Hidden = True
        workSheetRecap.Cells(0, SCOL_Recap + colST_CVG).EntireColumn.Hidden = True

        ' Hide the old ST columns

        For STi As Integer = 0 To xls_STs.Count
            workSheetRecap.Cells(0, SCOL_Recap + COLS_Recap + STi).EntireColumn.Hidden = True
        Next


        If xls_STs.Count > 0 Then
            For Each CHECKBOOK_GRP As String In xls_STs.Keys
                Dim CBI_first As Integer = -1

                Dim SLSf As String = ""
                Dim HRSf As String = ""

                For Each CHECKBOOK As String In xls_STs(CHECKBOOK_GRP)
                    Dim rowSPTCWRXC As DataRow = dst.Tables("SPTCWRXC").Rows.Find(CHECKBOOK)
                    Dim CBI As Integer = Val(rowSPTCWRXC.Item("CBI") & "")
                    If CBI_first = -1 Then
                        CBI_first = CBI
                        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_SLS).Value = CHECKBOOK_GRP
                    Else
                        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + 10).EntireColumn.Hidden = True
                        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_SLS).EntireColumn.Hidden = True
                        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_HRS).EntireColumn.Hidden = True
                        workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colST_CVG).EntireColumn.Hidden = True
                    End If


                    SLSf &= "+" & Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colBUDRTL)
                    HRSf &= "+" & Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI) * COLS_Recap + colBUDHRS)

                Next

                'workSheetRecap.Cells(0, SCOL_Recap + COLS_ST + (CBi) * COLS_Recap + 10).EntireColumn.ColumnWidth = 2 ' not nec
                Dim SLS_RC As String = Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_SLS)
                Dim HRS_RC As String = Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_HRS)
                Mid(SLSf, 1, 1) = "="
                Mid(HRSf, 1, 1) = "="
                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_SLS).Value = SLSf
                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_SLS).NumberFormat = "#,##0"
                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_HRS).Value = HRSf
                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_HRS).NumberFormat = "#,##0"

                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG).Formula = "=IF(" & SLS_RC & "=0,0,+" & HRS_RC & "*" & "100" & "/" & SLS_RC & ")"

                workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG).NumberFormat = "#,##0.0%"

                ' Paste the Formula to all of the weeks in the Weekly Section
                rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_SLS,
                                                     SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG)
                rangePaste_To = workSheetRecap.Range(SROW_Weeks + 2, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_SLS,
                                                     SROW_Weeks + Number_of_Weeks_in_Season,
                                                     SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG)
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.FormulasAndNumberFormats, SpreadsheetGear.PasteOperation.None, False, False)

                For iCVG As Integer = 1 To Number_of_Weeks_in_Season
                    Dim F As String = workSheetRecap.Cells(SROW_Weeks + iCVG, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG).Formula
                    F = Replace(F, "*100", "*" & CVG_PCT(iCVG))
                    workSheetRecap.Cells(SROW_Weeks + iCVG, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG).Formula = F

                    If iCVG = 1 Then ' COPY FORMULAS FROM WEEK 1 TO 6 MONTHS AND TOTAL LINE

                        'rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + iCVG - 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 2, _
                        '                                     SROW_Weeks + iCVG - 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0)
                        'rangePaste_To = workSheetRecap.Range(SROW_Weeks + iCVG - 11, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 2, _
                        '                                     SROW_Weeks + iCVG - 11, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0)
                        'rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.FormulasAndNumberFormats, SpreadsheetGear.PasteOperation.None, False, False)

                        rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + iCVG, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 2,
                                                             SROW_Weeks + iCVG, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0)
                        rangePaste_To = workSheetRecap.Range(SROW_Weeks + iCVG - 10, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 2,
                                                             SROW_Weeks + iCVG - 4, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0)
                        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.FormulasAndNumberFormats, SpreadsheetGear.PasteOperation.None, False, False)
                        workSheetRecap.Range(SROW_Weeks + iCVG - 10, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0,
                                             SROW_Weeks + iCVG - 4, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG - 0).NumberFormat = "#,##0.0%"
                    End If
                Next





                ' don't know why paste formats did not work for %
                range = workSheetRecap.Range(SROW_Weeks + 1, SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG,
                                             SROW_Weeks + Number_of_Weeks_in_Season,
                                             SCOL_Recap + COLS_ST + (CBI_first) * COLS_Recap + colST_CVG)
                range.NumberFormat = "#,##0.0%"

                rangePaste_To.Locked = True
                rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
            Next
        End If
        AutoFit_Summary_Columns()

    End Sub

    Sub XLS_Refresh_Stores()

        Dim worksheetSummary As SpreadsheetGear.IWorksheet = workbook.Worksheets("Summary")
        Dim iStore As Integer = 0
        Dim XLA As String = ""
        xls_STOREs.Clear()
        LOCK_SHEETS.Clear()

        Dim rows() As DataRow
        If EntryMode = "R" Then
            rows = ASCDATA1.SelectDistinct("ARTCUST2", New String() {"REGION_CODE", "SELL_CODE", "SELL_NAME"}).Select("", "REGION_CODE,SELL_CODE")
        Else
            rows = dst.Tables("ARTCUST2").Select("", "HAS_BUDGET DESC,CUST_CODE,CUST_STORE_NO") ' ASCDATA1.GetDataTable.Select("", "CUST_CODE,CUST_STORE_NO")
        End If

        Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")

        For Each row As DataRow In rows

            Dim C_VALUE As String
            Dim S_VALUE As String
            Dim CS As String
            Dim CS_NAME As String

            If EntryMode = "R" Then
                C_VALUE = row.Item("REGION_CODE")
                S_VALUE = row.Item("SELL_CODE") & "" ' temp
                CS = C_VALUE & "-" & S_VALUE
                CS_NAME = row.Item("SELL_NAME")
            Else
                C_VALUE = row.Item("CUST_CODE")
                S_VALUE = row.Item("CUST_STORE_NO")
                CS = C_VALUE & "-" & S_VALUE
                CS_NAME = row.Item("CUST_STORE_NAME")
            End If

            If Not xls_STOREs.Contains(CS) Then

                worksheet = workbook.Worksheets.Add
                worksheet.Name = CS

                rangeCopyFrom = worksheet2.Cells(Excel_Cell(0, 1) & ":" &
                                                 Excel_Cell(0, SCOL_Recap + COLS_ST + (CBi + 1) * COLS_Recap))
                rangePaste_To = worksheet.Cells(0, 0)
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

                iStore += 1

                worksheet.Cells(2, 3).Value = CS_NAME
                worksheet.Cells(3, 1).Value = C_VALUE
                worksheet.Cells(4, 1).Value = S_VALUE

                worksheet.Cells(SROW_Weeks + 1, SCOL_Recap + 1 + 1).Activate()
                worksheet.WindowInfo.FreezePanes = True

                worksheetSummary.Cells(5 + iStore, 2).Value = C_VALUE
                worksheetSummary.Cells(5 + iStore, 3).Value = S_VALUE
                worksheetSummary.Cells(5 + iStore, 4).Value = CS_NAME

                worksheetSummary.Hyperlinks.Add(worksheetSummary.Cells(5 + iStore, 4),
                                          "",
                                          "'" & CS & "'!D3",
                                         "Click Here to Navigate to " & CS,
                                         CS_NAME)
                ' worksheetSummary.Cells(5 + iStore, 4).Locked = False

                worksheetSummary.Cells(5 + iStore, 4 + 1).Formula = "='" & CS & "'!" & Excel_Cell0(3, SCOL_Recap + colBUDRTL)

                For CBi2 As Integer = 0 To CBi
                    Dim COLS_STi As Integer = COLS_ST * System.Math.Sign(CBi2)
                    worksheetSummary.Cells(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 1).Formula = "='" & CS & "'!" & Excel_Cell0(3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD)
                    worksheetSummary.Cells(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 2).Formula = "='" & CS & "'!" & Excel_Cell0(4, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD)
                    worksheetSummary.Cells(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 3).Formula = "='" & CS & "'!" & Excel_Cell0(14, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD)
                    worksheetSummary.Cells(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 4).Formula = "=" & Excel_Cell0(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 3) & "-" & Excel_Cell0(5 + iStore, 4 + 1 + (COLS_Summary) * CBi2 + 2)
                Next

                xls_STOREs.Add(CS)
                XLA &= "+'" & CS & "'!" & Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + colLYRHRS)
                '    Excel_Cell0(SROW_Weeks + 1, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colLYRHRS)

                If ASCMAIN1.Running_in_VS AndAlso S_VALUE = "000495" And C_VALUE = "MACYS" Then Stop

                If EntryMode <> "R" Then
                    For CBi2 As Integer = 1 To CBi
                        Dim rowSPTCWRXC As DataRow = dst.Tables("SPTCWRXC").Select("CBI = " & CStr(CBi2))(0)
                        Dim CHECKBOOK As String = rowSPTCWRXC.Item("CHECKBOOK")

                        Dim sqlw As String = "CUST_CODE = '" & C_VALUE & "' and CHECKBOOK = '" & CHECKBOOK & "'"
                        Dim PROMOs As String = ""
                        For Each rowSPTCOOPX As DataRow In dst.Tables("SPTCOOPX").Select(sqlw, "EXPENSE_TYPE_CODE")

                            Dim OPS_YYYYWW As String = rowSPTCOOPX.Item("OPS_YYYYWW") & ""
                            Dim BOOKING_NAME As String = rowSPTCOOPX.Item("BOOKING_NAME") & ""
                            Dim EXPENSE_TYPE_CODE As String = rowSPTCOOPX.Item("EXPENSE_TYPE_CODE") & ""
                            Dim EVENT_GROUP_NO As String = rowSPTCOOPX.Item("EVENT_GROUP_NO") & ""
                            Dim VEHICLE_CODE As String = rowSPTCOOPX.Item("VEHICLE_CODE") & ""
                            Dim rowGLTPARM3 As DataRow = dst.Tables("GLTPARM3").Rows.Find(OPS_YYYYWW)

                            If ASCMAIN1.Running_in_VS AndAlso S_VALUE = "000495" And C_VALUE = "MACYS" AndAlso BOOKING_NAME.StartsWith("MAY WK3 SECONDARY") Then Stop


                            Dim skip_event As Boolean = False
                            ' If EXPENSE_TYPE_CODE = "RTLEVENTS" And EVENT_GROUP_NO <> "" Then
                            'skip_event = True
                            ' sarah says hold off for now with RTLEVENTS
                            ' End If

                            If EntryMode = "R" Then
                                ' WE MAY NEED TO SEE IF THE AE HAS ANY STORES THAT HAVE THIS EVENT

                            Else

                                If dst.Tables("SPTSFOC9").Select("EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'").Length <> 0 Then
                                    If dst.Tables("SPTSFOC9").Rows.Find(New String() {EVENT_GROUP_NO, C_VALUE, S_VALUE}) Is Nothing Then
                                        skip_event = True ' THIS EVENT GROUP IS BY STORE AND IS NOT FOR THIS CUST-STORE
                                    End If
                                End If
                            End If


                            If skip_event Then

                            Else

                                If rowGLTPARM3 IsNot Nothing Then
                                    Dim WEEK_NO As Integer = Val(rowGLTPARM3.Item("WEEK_NO") & "")
                                    Dim PNOTE As String = worksheet.Cells(SROW_Weeks + WEEK_NO, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colPNOTES).Value & ""
                                    If PNOTE <> "" Then PNOTE &= ";"
                                    If EXPENSE_TYPE_CODE = "RTLEVENTS" And EVENT_GROUP_NO <> "" Then
                                        PNOTE &= VEHICLE_CODE & ":" & BOOKING_NAME
                                    Else
                                        PNOTE &= Mid(EXPENSE_TYPE_CODE, 1, 3) & ":" & BOOKING_NAME
                                    End If
                                    worksheet.Cells(SROW_Weeks + WEEK_NO, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + colPNOTES).Value = PNOTE
                                End If

                            End If
                        Next
                    Next

                    If Not isAC And Not load_from_AC AndAlso row.Item("SELL_CODE_AC") & "" <> "" Then
                        'If Not isAC AndAlso row.Item("SELL_CODE_AC") & "" <> "" Then
                        LOCK_SHEETS.Add(CS, row.Item("SELL_CODE_AC"))
                        'worksheet.ProtectContents = True
                        'worksheet.Protect(XLS_PWD)
                    End If

                End If
            End If
        Next

        Dim workSheetRecap As SpreadsheetGear.IWorksheet = workbook.Worksheets("Recap")

        ' Lock down the Notes section of the Recap Sheet (Month Rows) after having copied the Recap Sheet to all stores
        range = workSheetRecap.Cells(SROW_Months + 1, 6, SROW_Months + 6, 6)
        range.Locked = True
        range.Interior.Color = SpreadsheetGear.Colors.White

        ' Lock down the Notes section of the Recap Sheet (Week Rows) after having copied the Recap Sheet to all stores
        range = workSheetRecap.Cells(SROW_Weeks + 1, 6,
                                     SROW_Weeks + Number_of_Weeks_in_Season, 6)
        range.Locked = True
        range.Interior.Color = SpreadsheetGear.Colors.White

        ' Set the Formula for Total Hours LY as the sum of each Stores Total Hours LY
        ' this formula will be copied to several ranges
        ' for each range we paste to, we will lock down the range for editing
        workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + colLYRHRS).Formula = "=" & XLA
        rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + colLYRHRS)

        For CBi2 As Integer = 0 To CBi
            Dim COLS_STi As Integer = COLS_ST * System.Math.Sign(CBi2)

            ' Paste the Formula to the Sales Goal
            rangePaste_To = workSheetRecap.Range(SROW_Months - 4, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDRTL,
                                                 SROW_Months - 3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDRTL)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

            ' Paste the Formula to the Spend Budget
            rangePaste_To = workSheetRecap.Range(SROW_Months - 4, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD,
                                                 SROW_Months - 3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

            ' No need to paste to Plan Rate - just lock it down
            rangePaste_To = workSheetRecap.Range(SROW_Months - 3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colPLANRATE)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
            'rangePaste_To.Formula = "=IFERROR(" & Excel_Cell0(4, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDHRS) & "/" & _
            '                                      Excel_Cell0(4, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD) & ",0)"
            ' shouldn't the above formula be reversed?
            '=IFERROR(O5/P5,0)

            ' Paste the Formula to the Hours to Plan
            rangePaste_To = workSheetRecap.Range(SROW_Months - 3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colHOURS2PLAN)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

            ' Paste the Formula to all of the weeks in the Weekly Section
            rangePaste_To = workSheetRecap.Range(SROW_Weeks + 1, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colLYRHRS,
                                                 SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colLYRHRS)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
            rangePaste_To.Locked = True
            rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White

            ' Paste the Formula to Spend LY, Sales LY, Sales TY, Hours TY, and Spend TY
            ' rangeCopyFrom = workSheetRecap.Cells(SROW_Weeks + 1, SCOL_Recap + 1, SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + 1)
            For Each c As Integer In New Integer() {colLYRACT, colLYRSLS, colBUDRTL, colBUDHRS, colBUDMOD}
                rangePaste_To = workSheetRecap.Range(SROW_Weeks + 1, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + c,
                                                     SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + c)
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                rangePaste_To.Locked = True
                rangePaste_To.Interior.Color = SpreadsheetGear.Colors.White
            Next

            Dim XC As Integer = 4 + 1 + (COLS_Summary) * CBi2
            If CBi2 = 0 Then worksheetSummary.Cells(3, XC).Formula = "=SUM(" & Excel_Cell0(5 + 1, XC) & ":" & Excel_Cell0(5 + xls_STOREs.Count, XC) & ")"
            XC += 1 : worksheetSummary.Cells(3, XC).Formula = "=SUM(" & Excel_Cell0(5 + 1, XC) & ":" & Excel_Cell0(5 + xls_STOREs.Count, XC) & ")"
            XC += 1 : worksheetSummary.Cells(3, XC).Formula = "=SUM(" & Excel_Cell0(5 + 1, XC) & ":" & Excel_Cell0(5 + xls_STOREs.Count, XC) & ")"
            XC += 1 : worksheetSummary.Cells(3, XC).Formula = "=SUM(" & Excel_Cell0(5 + 1, XC) & ":" & Excel_Cell0(5 + xls_STOREs.Count, XC) & ")"

            range = worksheetSummary.Cells(5 + 1, XC - 2, 5 + xls_STOREs.Count, XC - 1 + COLS_Summary - 3)
            range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous

        Next

        'If EntryMode = "R" Then
        'Else
        For CBi2 As Integer = 0 To CBi
            Dim COLS_STi As Integer = COLS_ST * System.Math.Sign(CBi2)
            ' workSheetRecap.Range(SROW_Weeks + 1, SCOL_Recap + (CBi2 * COLS_Recap) + c, SROW_Weeks + Number_of_Weeks_in_Season, SCOL_Recap + (CBi2 * COLS_Recap) + c)
            worksheetSummary.Cells(1, 4 + 1 + (COLS_Summary) * CBi2 + 3).Formula = "='Recap'!" & Excel_Cell0(3, SCOL_Recap + COLS_STi + (CBi2 * COLS_Recap) + colBUDMOD)
        Next
        'End If

        If EntryMode = "R" Then
            Dim RX As Integer = 5
            Dim C_VALUE_sum As String = ""

            Dim F As String = ""

            For I As Integer = 1 To xls_STOREs.Count
                RX += 1
                If worksheetSummary.Cells(RX, 2).Value <> C_VALUE_sum Then
                    C_VALUE_sum = worksheetSummary.Cells(RX, 2).Value
                    worksheetSummary.Cells(RX, 2).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                    RX += 1
                    worksheetSummary.Cells(RX - 1, 2).Value = C_VALUE_sum

                    Dim C_VALUE_desc As String = ""
                    Dim rowSOTSREG1 As DataRow = dst.Tables("SOTSREG1").Rows.Find(C_VALUE_sum)
                    If rowSOTSREG1 IsNot Nothing Then
                        C_VALUE_desc = rowSOTSREG1.Item("REGION_DESC") & ""
                    End If
                    worksheetSummary.Cells(RX - 1, 4).Value = C_VALUE_desc

                    Dim C As Integer = 0
                    Do While worksheetSummary.Cells(RX + C + 1, 2).Value & "" = C_VALUE_sum
                        C += 1
                    Loop
                    worksheetSummary.Cells(RX - 1, 5).Formula = "=SUM(" & Excel_Cell0(RX, 5) & ":" & Excel_Cell0(RX + C, 5) & ")"
                    worksheetSummary.Cells(RX - 1, 5).EntireRow.Interior.Color = SpreadsheetGear.Colors.Aqua
                    rangeCopyFrom = worksheetSummary.Cells(RX - 1, 5)
                    F &= "+" & Excel_Cell0(RX - 1, 5)

                    For CBi2 As Integer = 0 To CBi
                        For Each ix As Integer In New Integer() {0, 1, 2, 3}
                            Dim Cx As Integer = SCOL_Recap + (CBi2 * COLS_Summary) + ix + 1
                            rangePaste_To = worksheetSummary.Range(RX - 1, Cx, RX - 1, Cx)
                            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                        Next
                    Next
                End If
            Next

            worksheetSummary.Cells(3, 5).Formula = "=" & Mid(F, 2)
            rangeCopyFrom = worksheetSummary.Cells(3, 5)

            For CBi2 As Integer = 0 To CBi
                For Each ix As Integer In New Integer() {0, 1, 2, 3}
                    Dim Cx As Integer = SCOL_Recap + (CBi2 * COLS_Summary) + ix + 1
                    rangePaste_To = worksheetSummary.Range(3, Cx, 3, Cx)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                Next
            Next
        End If



    End Sub

    Sub Get_Budgets()

        Dim sqlSum As String = ", Sum (BUDGET_MOD) BUDGET_MOD, Sum (BUDGET_RTL) BUDGET_RTL"

        Dim between_TY As String = "'" & YPs(1) & "' and '" & YPs(6) & "'"

        Dim sqlGBy As String = "CUST_CODE, CUST_STORE_NO"
        If (EntryMode = "R") Then
            'sqlGBy = "REGION_CODE, SELL_CODE"
            sqlGBy = "REGION_CODE, NVL(SELL_CODE,'000')"
        End If

        ASCMAIN1.sql = "" _
            & "Select " & sqlGBy & ", CHECKBOOK" & vbCrLf _
            & sqlSum & vbCrLf _
            & "from (" & vbCrLf _
            & "Select SPTMBUD1.CUST_CODE, SPTMBUD1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, SPTMBUD1.OPS_YYYYPP, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ", Sum (SPTMBUD1.BUDGET) BUDGET_MOD, 0 BUDGET_RTL" & vbCrLf _
            & "from SPTMBUD1," & ARTCUST2 & " ARTCUST2," & SPTCWRXW & " SPTCWRXW" & vbCrLf _
            & "where SPTMBUD1.OPS_YYYYPP between " & between_TY & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = SPTMBUD1.CUST_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_STORE_NO = SPTMBUD1.CUST_STORE_NO" & vbCrLf _
            & "  and SPTCWRXW.COLLECTION_CODE = SPTMBUD1.COLLECTION_CODE" & vbCrLf _
            & "group by SPTMBUD1.CUST_CODE, SPTMBUD1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, SPTMBUD1.OPS_YYYYPP, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & "union" & vbCrLf _
            & "Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, RSTBUDR1.OPS_YYYYPP, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ", 0 BUDGET_MOD, Sum (RSTBUDR1.BUDGET) BUDGET_RTL" & vbCrLf _
            & "from RSTBUDR1," & ARTCUST2 & " ARTCUST2," & SPTCWRXW & " SPTCWRXW" & vbCrLf _
            & "where RSTBUDR1.OPS_YYYYPP between " & between_TY & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = RSTBUDR1.CUST_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_STORE_NO = RSTBUDR1.CUST_STORE_NO" & vbCrLf _
            & "  and SPTCWRXW.COLLECTION_CODE = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & "group by RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, RSTBUDR1.OPS_YYYYPP, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ") group by " & sqlGBy & ", CHECKBOOK"

        Dim tbl_Budgets As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "BUD", 3)

        WorkbookView1.GetLock()

        For Each CS As String In xls_STOREs
            Dim C_VALUE As String = Split(CS, "-")(0)
            Dim S_VALUE As String = Split(CS, "-")(1)

            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(CS)
            CBi = 0
            For Each rowSPTCWRXC As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
                CBi += 1
                Dim CHECKBOOK As String = rowSPTCWRXC.Item("CHECKBOOK")
                Dim C0 As Integer = SCOL_Recap + COLS_ST + CBi * COLS_Recap
                Dim rowBudget As DataRow = tbl_Budgets.Rows.Find(New String() {C_VALUE, S_VALUE, CHECKBOOK})
                '   If ASCMAIN1.Running_in_VS And C_VALUE = "BERGDORF" And S_VALUE = "000063" And CHECKBOOK = "VCA" Then Stop
                Dim BUDGET_MOD As Decimal = 0
                Dim BUDGET_RTL As Decimal = 0
                ' If C_VALUE = "DILLARDS" And S_VALUE = "000907" Then STOP
                If rowBudget IsNot Nothing Then
                    BUDGET_MOD = Val(rowBudget.Item("BUDGET_MOD") & "")
                    BUDGET_RTL = Val(rowBudget.Item("BUDGET_RTL") & "")
                End If
                ws.Cells(3, C0 + colBUDMOD).Value = BUDGET_MOD
                ws.Cells(3, C0 + colBUDRTL).Value = BUDGET_RTL

                If EntryMode = "N" Then
                    ws.Cells(4, C0 + colBUDMOD).Value = BUDGET_MOD
                    ws.Cells(4, C0 + colBUDRTL).Value = BUDGET_RTL
                End If
            Next
        Next

        WorkbookView1.ReleaseLock()
    End Sub

    Sub Get_Actuals()

        Dim sqlSum As String = ""
        Dim CBi As Integer = 0
        For Each row As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
            CBi += 1
            Dim CHECKBOOK As String = row.Item("CHECKBOOK")
            For Each C As String In New String() {"HRS", "AMT", "RTL"}
                sqlSum &= ", Sum (Decode(CHECKBOOK,'" & CHECKBOOK & "'," & C & ",0)) " & C & "_CB" & Format(CBi, "00")
            Next
        Next

        Dim between_LY As String = "'" & LYWKs(1) & "' and '" & LYWKs(Number_of_Weeks_in_Season) & "'"
        Dim between_TY As String = "'" & TYWKs(1) & "' and '" & TYWKs(AYWi) & "'"

        Dim sqlGBy As String = "CUST_CODE, CUST_STORE_NO"
        If (EntryMode = "R") Then
            'sqlGBy = "REGION_CODE, SELL_CODE"
            sqlGBy = "REGION_CODE, NVL(SELL_CODE,'000')"
        End If

        ASCMAIN1.sql = "" _
            & "Select " & sqlGBy & ", OPS_YYYYWW" & vbCrLf _
            & sqlSum & vbCrLf _
            & "from (" & vbCrLf _
            & "Select SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, SPTCWRX2.OPS_YYYYWW, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ", Sum (SPTCWRX2.BILL_HOURS) HRS, Sum (SPTCWRX2.BILL_AMT) AMT, 0 RTL" & vbCrLf _
            & "from SPTCWRX2," & ARTCUST2 & " ARTCUST2," & SPTCWRXW & " SPTCWRXW" & vbCrLf _
            & "where SPTCWRX2.OPS_YYYYWW between " & between_LY & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = SPTCWRX2.CUST_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_STORE_NO = SPTCWRX2.CUST_STORE_NO" & vbCrLf _
            & "  and SPTCWRXW.COLLECTION_CODE = SPTCWRX2.COLLECTION_CODE" & vbCrLf _
            & "group by SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, SPTCWRX2.OPS_YYYYWW, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & "union" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, RSTRETL1.OPS_YYYYWW, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ", 0 HRS, 0 AMT, Sum (RSTRETL1.AMT_SOLD) RTL" & vbCrLf _
            & "from RSTRETL1,ICTITEM1," & ARTCUST2 & " ARTCUST2," & SPTCWRXW & " SPTCWRXW" & vbCrLf _
            & "where RSTRETL1.OPS_YYYYWW between " & between_LY & vbCrLf _
            & "  and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_CODE = RSTRETL1.CUST_CODE" & vbCrLf _
            & "  and ARTCUST2.CUST_STORE_NO = RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & "  and SPTCWRXW.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK, RSTRETL1.OPS_YYYYWW, ARTCUST2.REGION_CODE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ") group by " & sqlGBy & ", OPS_YYYYWW"

        Dim tbl_LY As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "LY", 3)

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, between_LY, between_TY)
        Dim tbl_TY As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "LY", 3)

        ' This method should be called only once, when EntryMode = "N"

        WorkbookView1.GetLock()

        For Each CS As String In xls_STOREs
            Dim C_VALUE As String = Split(CS, "-")(0)
            Dim S_VALUE As String = Split(CS, "-")(1)

            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(CS)

            CBi = 0
            For Each rowSPTCWRXC As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
                CBi += 1
                Dim CHECKBOOK As String = rowSPTCWRXC.Item("CHECKBOOK")
                Dim rowSPTMXWS4 As DataRow = dst.Tables("SPTMXWS4").Rows.Find(New Object() {SEASON_CODE, C_VALUE, S_VALUE, CHECKBOOK})
                If rowSPTMXWS4 IsNot Nothing Then
                    Dim BUDGET_RTL As Decimal = Val(rowSPTMXWS4.Item("BUDGET_RTL") & "")
                    Dim BUDGET_MOD As Decimal = Val(rowSPTMXWS4.Item("BUDGET_MOD") & "")
                    Dim BUDGET_RTL_REV As Decimal = Val(rowSPTMXWS4.Item("BUDGET_RTL_REV") & "")
                    Dim BUDGET_MOD_REV As Decimal = Val(rowSPTMXWS4.Item("BUDGET_MOD_REV") & "")
                    Dim BUDGET_RATE As Decimal = Val(rowSPTMXWS4.Item("BUDGET_RATE") & "")

                    ' If C_VALUE = "DILLARDS" And S_VALUE = "000907" Then Stop

                    ''' SP SAYS DO NOT DO THIS ws.Cells(3, SCOL_Recap + CBi * COLS_Recap + 4).Value = BUDGET_RTL
                    ''' SP SAYS DO NOT DO THIS ws.Cells(3, SCOL_Recap + CBi * COLS_Recap + 7).Value = BUDGET_MOD
                    'If BUDGET_RTL_REV <> 0 And BUDGET_RTL_REV <> BUDGET_RTL Then
                    ws.Cells(4, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colBUDRTL).Value = BUDGET_RTL_REV
                    'End If
                    'If BUDGET_MOD_REV <> 0 And BUDGET_MOD_REV <> BUDGET_MOD Then
                    ws.Cells(4, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colBUDMOD).Value = BUDGET_MOD_REV
                    'End If
                    If EntryMode = "R" Then
                        Dim SPEND As String = Excel_Cell0(14, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colBUDMOD)
                        Dim HOURS As String = Excel_Cell0(14, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colBUDHRS)
                        Dim F As String = "=IFERROR(" & SPEND & "/" & HOURS & ",20)"

                        ws.Cells(4, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPLANRATE).Formula = F
                    Else
                        If BUDGET_RATE <> 0 Then
                            ws.Cells(4, SCOL_Recap + COLS_ST + CBi * COLS_Recap + colPLANRATE).Value = BUDGET_RATE
                        End If
                    End If
                End If
            Next

            For I As Integer = 1 To 6
                Dim rowSPTMXWS3 As DataRow = dst.Tables("SPTMXWS3").Rows.Find(New Object() {SEASON_CODE, C_VALUE, S_VALUE, I, 0})
                If rowSPTMXWS3 IsNot Nothing Then
                    ws.Cells(SROW_Months + I, 6).Value = rowSPTMXWS3.Item("NOTES")
                End If
            Next

            Dim AYWi_Spend As Integer = AYWi
            If YW_LAST_CWRX < AYW Then AYWi_Spend -= 1

            For Each rowGLTPARM3 As DataRow In dst.Tables("GLTPARM3").Select("", "YYYYWW")
                Dim WEEK_NO As Integer = Val(rowGLTPARM3.Item("WEEK_NO") & "")
                Dim MONTH_NO As Integer = Val(rowGLTPARM3.Item("MONTH_NO") & "")

                Dim LEGEND As String = rowGLTPARM3.Item("LEGEND")
                Dim WEEK_END_DATE As Date = rowGLTPARM3.Item("WEEK_END_DATE")
                Dim TYW As String = TYWKs(WEEK_NO)
                Dim LYW As String = LYWKs(WEEK_NO)
                Dim rowLY As DataRow = tbl_LY.Rows.Find(New String() {C_VALUE, S_VALUE, LYW})
                Dim rowTY As DataRow = tbl_TY.Rows.Find(New String() {C_VALUE, S_VALUE, TYW})

                Dim rowSPTMXWS3 As DataRow = dst.Tables("SPTMXWS3").Rows.Find(New Object() {SEASON_CODE, C_VALUE, S_VALUE, MONTH_NO, WEEK_NO})
                If rowSPTMXWS3 IsNot Nothing Then
                    ws.Cells(SROW_Weeks + WEEK_NO, 6).Value = rowSPTMXWS3.Item("NOTES")
                End If

                CBi = 0
                For Each rowSPTCWRXC As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
                    CBi += 1
                    Dim CHECKBOOK As String = rowSPTCWRXC.Item("CHECKBOOK")

                    ' Mens or Womens?  Use Checkbook Gender, unless U, in which case use the Key Collection
                    Dim COLLECTION_GENDER As String = rowSPTCWRXC.Item("COLLECTION_GENDER") & ""
                    If COLLECTION_GENDER <> "M" And COLLECTION_GENDER <> "W" Then
                        Dim COLLECTION_CODE As String = rowSPTCWRXC.Item("COLLECTION_CODE")
                        'Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                        Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                        COLLECTION_GENDER = rowICTCOLL1.Item("COLLECTION_GENDER") & ""
                    End If

                    Dim MWC As Integer = 6 - 1 ' F
                    If COLLECTION_GENDER = "M" Then
                        MWC = 5 - 1 ' E
                    End If

                    Dim C0 As Integer = SCOL_Recap + COLS_ST + CBi * COLS_Recap
                    Dim HRS As Decimal = 0
                    Dim AMT As Decimal = 0
                    Dim RTL As Decimal = 0

                    ' LY Actuals

                    If rowLY Is Nothing Then
                        HRS = 0
                        AMT = 0
                        RTL = 0
                    Else
                        HRS = Val(rowLY.Item("HRS_CB" & Format(CBi, "00")) & "")
                        AMT = Val(rowLY.Item("AMT_CB" & Format(CBi, "00")) & "")
                        RTL = Val(rowLY.Item("RTL_CB" & Format(CBi, "00")) & "")
                    End If
                    ws.Cells(SROW_Weeks + WEEK_NO, C0 + colLYRHRS).Value = HRS
                    ws.Cells(SROW_Weeks + WEEK_NO, C0 + colLYRACT).Value = AMT
                    ws.Cells(SROW_Weeks + WEEK_NO, C0 + colLYRSLS).Value = RTL

                    'If EntryMode = "N" Then
                    ' NOT SURE WHETHER WE SHOULD HANDLE THIS FOR ALL ENTRYMODES
                    'ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDRTL - 1).Formula = "=" & Excel_Cell0(4, C0 + colBUDRTL) & "*" & Excel_Cell0(SROW_Weeks + WEEK_NO, MWC)
                    ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDRTL).Formula = "=" & Excel_Cell0(4, C0 + colBUDRTL) & "*" & Excel_Cell0(SROW_Weeks + WEEK_NO, MWC)
                    'End If

                    ' TY Actuals

                    HRS = 0
                    AMT = 0
                    RTL = 0
                    If ASCMAIN1.Running_in_VS Then
                        '  If C_VALUE = "NEIMAN" And S_VALUE = "001005" And CHECKBOOK = "VCA" And (WEEK_NO >= 9 And WEEK_NO <= 12) Then Stop
                    End If
                    Dim rowSPTMXWS2 As DataRow = dst.Tables("SPTMXWS2").Rows.Find(New Object() {SEASON_CODE, C_VALUE, S_VALUE, CHECKBOOK, MONTH_NO, WEEK_NO})
                    If rowSPTMXWS2 IsNot Nothing Then
                        RTL = Val(rowSPTMXWS2.Item("TY_SALES") & "")
                        HRS = Val(rowSPTMXWS2.Item("TY_HOURS") & "")
                        AMT = Val(rowSPTMXWS2.Item("TY_SPEND") & "")
                    End If

                    If WEEK_NO <= AYWi Then
                        If rowTY IsNot Nothing Then
                            ' If WEEK_NO < AYWi Then ' OR TOTALLY ACTUALIZED - LIKE RUNNING A FINAL REPORT - NEED TO SET A FLAG
                            If WEEK_NO <= AYWi_Spend Then
                                HRS = Val(rowTY.Item("HRS_CB" & Format(CBi, "00")) & "")
                                AMT = Val(rowTY.Item("AMT_CB" & Format(CBi, "00")) & "")
                            End If
                            RTL = Val(rowTY.Item("RTL_CB" & Format(CBi, "00")) & "")
                        Else
                            If WEEK_NO < AYWi Then
                                HRS = 0
                                AMT = 0
                                ' RTL = 0
                            End If
                            RTL = 0
                        End If
                        ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDHRS).Value = HRS
                        ' If WEEK_NO < AYWi Then
                        If WEEK_NO <= AYWi_Spend Then
                            ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDMOD).Value = AMT
                        Else
                            If EntryMode = "R" Then
                                ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDMOD).Value = AMT
                            Else
                                ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDMOD).Formula = "=" & Excel_Cell0(4, C0 + colPLANRATE) & "*" & Excel_Cell0(SROW_Weeks + WEEK_NO, C0 + colBUDHRS)
                            End If
                        End If
                        ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDRTL).Value = RTL
                    Else
                        If rowSPTMXWS2 IsNot Nothing Then
                            If blnConsolidated Or (EntryMode = "R") Or (SEASON_SLS_BUD_LOCKED = "1") Then ' THIS IS WHERE WE REPLACE THE FORMULA FOR RTL WITH A VALUE
                                ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDRTL).Value = RTL
                            End If
                            If blnConsolidated Or (EntryMode = "R") Then
                                ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDMOD).Value = AMT
                            End If
                            ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDHRS).Value = HRS
                        End If
                        If blnConsolidated Or (EntryMode = "R") Then
                        Else
                            ws.Cells(SROW_Weeks + WEEK_NO, C0 + colBUDMOD).Formula = "=" & Excel_Cell0(4, C0 + colPLANRATE) & "*" & Excel_Cell0(SROW_Weeks + WEEK_NO, C0 + colBUDHRS)
                        End If
                    End If
                Next
            Next

            If AYWi > 0 Then
                'Dim AYWi_Spend As Integer = AYWi
                'If YW_LAST_CWRX < AYW Then AYWi_Spend -= 1

                CBi = 0
                For Each rowSPTCWRXC As DataRow In dst.Tables("SPTCWRXC").Select("", "CHECKBOOK_SEQ,CHECKBOOK")
                    CBi += 1
                    Dim C0 As Integer = SCOL_Recap + COLS_ST + CBi * COLS_Recap

                    range = ws.Range(SROW_Weeks + 1, C0 + colBUDRTL,
                                     SROW_Weeks + AYWi, C0 + colBUDRTL)
                    range.Locked = True
                    range.Interior.Color = SpreadsheetGear.Colors.White

                    If C_VALUE = "IPLBAE" Then
                        range = ws.Range(SROW_Weeks + 1, C0 + colBUDRTL, SROW_Weeks + Number_of_Weeks_in_Season, C0 + colBUDRTL)
                        range.Locked = True
                        range.Interior.Color = SpreadsheetGear.Colors.White

                        range = ws.Range(SROW_Weeks + 1, C0 + colBUDHRS, SROW_Weeks + Number_of_Weeks_in_Season, C0 + colBUDHRS)
                        range.Locked = True
                        range.Interior.Color = SpreadsheetGear.Colors.White
                        range.Value = 0
                    ElseIf AYWi_Spend > 0 Then
                        range = ws.Range(SROW_Weeks + 1, C0 + colBUDHRS, SROW_Weeks + AYWi_Spend, C0 + colBUDHRS)
                        range.Locked = True
                        range.Interior.Color = SpreadsheetGear.Colors.White
                    End If
                Next
            End If
        Next

        WorkbookView1.ReleaseLock()

    End Sub

    Private Sub grdRSTSSPLX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTMXWSX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SELL_CODE").Text = e.Row.Cells("SELL_CODE").Value & ""
            ' Click_Command("View")
            Click_Command("Edit")
        End If
    End Sub

    Sub Prepare_Summary()

    End Sub

    Sub Prepare_Work_Tables()

        RSTBUDR1 = TAC.RSCMAIN1.RSTBUDR1_as_YP

        ASCMAIN1.sql = "Select Distinct ICTCOLL1.COLLECTION_CODE, SPTCWRXC.CHECKBOOK from ICTCOLL1,SPTCWRXC" & vbCrLf _
            & " where SPTCWRXC.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and (SPTCWRXC.COLLECTION_GENDER = 'U' or ICTCOLL1.COLLECTION_GENDER = SPTCWRXC.COLLECTION_GENDER)"

        ASCMAIN1.sql = "Select Distinct ICTCOLL1.COLLECTION_CODE, ICTCOLL0.CHECKBOOK from ICTCOLL1,ICTCOLL0" & vbCrLf _
             & " where ICTCOLL0.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
             & "   and ICTCOLL1.COLLECTION_STATUS = 'A'"

        SPTCWRXW = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SPTCWRXW & " Add Primary Key (COLLECTION_CODE)")

    End Sub

    Sub Create_ARTCUST2()

        Dim SYP As String = "ARTCUST2.SELL_CODE"
        If isAC Or load_from_AC Then
            SYP = "ARTCUST2.SELL_CODE_AC"
        End If
        If YPs IsNot Nothing Then
            Dim SELL_CODE_LAST_YP_NEXT As String = ASCMAIN1.Period_Calc(YPs(6), 1)
            Dim exp As String = ".SELL_CODE,CASE WHEN NVL(ARTCUST2.SELL_CODE_LAST_YP,'000000') >= '" & YPs(1) & "' THEN ARTCUST2.SELL_CODE_LAST ELSE NULL END)"
            If isAC Or load_from_AC Then
                exp = exp.Replace(".SELL_CODE", ".SELL_CODE_AC")
            End If
            SYP = "NVL(" & IIf(useFrozenAlignment, "ARTCUST7", "ARTCUST2") & exp
            'SYP = "NVL(ARTCUST2.SELL_CODE,CASE WHEN NVL(ARTCUST2.SELL_CODE_LAST_YP,'000000') BETWEEN '" & YPs(1) & "' AND '" & SELL_CODE_LAST_YP_NEXT & "' THEN ARTCUST2.SELL_CODE_LAST ELSE NULL END)"
            ' SYP = "NVL(ARTCUST2.SELL_CODE,CASE WHEN NVL(ARTCUST2.SELL_CODE_LAST_YP,'000000') BETWEEN '" & YPs(1) & "' AND '" & YPs(6) & "' THEN ARTCUST2.SELL_CODE_LAST ELSE NULL END)"
        End If

        ASCMAIN1.sql = "Select ARTCUST2.*" & vbCrLf _
            & ",SOTSELL1.SELL_NAME,SOTSELL1.REGION_CODE,SOTSREG1.REGION_DESC" & vbCrLf _
            & " from ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
            & IIf(useFrozenAlignment, ", ARTCUST7", "") & vbCrLf _
            & " where SOTSELL1.SELL_CODE = " & SYP & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
            & IIf(useFrozenAlignment, " AND ARTCUST2.CUST_CODE = ARTCUST7.CUST_CODE AND ARTCUST2.CUST_STORE_NO = ARTCUST7.CUST_STORE_NO AND ARTCUST7.OPS_YYYYPP = '" & frozenAlignment_YP & "'", "") & vbCrLf _
            & IIf(EntryMode = "R",
                  "   and SOTSREG1.REGION_CODE is Not Null" & IIf(SELL_CODES_consolidated.Count = 0, "", " and NVL(ARTCUST2.SELL_CODE,ARTCUST2.SELL_CODE_LAST)  in ('" & Join(SELL_CODES_consolidated.ToArray, "','") & "')" & vbCrLf),
                  "   and " & SYP & " = '" & SELL_CODE & "'")
        If ARTCUST2 = "" Then
            ARTCUST2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST2 & " Add Primary Key (CUST_CODE,CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCUST2)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCUST2 & " " & ASCMAIN1.sql)
        End If

        If useFrozenAlignment And EntryMode = "E" Then ' TO AVOID THINGS LIKE RECAP AND SUMMARIES WHICH MIGHT REQUIRE MORE ANALYSIS
            ASCDATA1.ExecuteSQL("Update " & ARTCUST2 & " Set SELL_CODE = '" & SELL_CODE & "' where SELL_CODE is Null")
            ' WAS CAUSING AN ERROR WHEN LOADING UP AE 140 ON MACYS-000171 WHICH HAD A NULL VALUE IN SELL_CODE
        End If

        ' ADDED  and CUST_STORE_STATUS = 'I' W/DM BECAUSE STORES WERE RE-ACTIVATED AND THE HISTORY IN THE LAST FIELDS WAS FORCING THE STORE TO THE HISTORICAL AE/AC
        ' REVERTED BECAUSE WE MADE A CHANGE TO CUST/STORE FM SO THAT THE HISTORY FIELDS ARE CLEARED WHEN A DOOR IS RE-ACTIVATED
        'ASCMAIN1.sql = "Update " & ARTCUST2 & " Set SELL_CODE = SELL_CODE_LAST where SELL_CODE is Null and SELL_CODE_LAST is Not Null and CUST_STORE_STATUS = 'I'"
        ASCMAIN1.sql = "Update " & ARTCUST2 & " Set SELL_CODE = SELL_CODE_LAST where SELL_CODE is Null and SELL_CODE_LAST is Not Null"
        ASCDATA1.ExecuteSQL()

        ' ADDED  and CUST_STORE_STATUS = 'I' W/DM BECAUSE STORES WERE RE-ACTIVATED AND THE HISTORY IN THE LAST FIELDS WAS FORCING THE STORE TO THE HISTORICAL AE/AC
        ' REVERTED BECAUSE WE MADE A CHANGE TO CUST/STORE FM SO THAT THE HISTORY FIELDS ARE CLEARED WHEN A DOOR IS RE-ACTIVATED
        'ASCMAIN1.sql = "Update " & ARTCUST2 & " Set SELL_CODE_AC = SELL_CODE_AC_LAST where SELL_CODE_AC is Null and SELL_CODE_AC_LAST is Not Null and CUST_STORE_STATUS = 'I'"
        ASCMAIN1.sql = "Update " & ARTCUST2 & " Set SELL_CODE_AC = SELL_CODE_AC_LAST where SELL_CODE_AC is Null and SELL_CODE_AC_LAST is Not Null"
        ASCDATA1.ExecuteSQL()

        If isAC Or load_from_AC Then
            ASCMAIN1.sql = "Update " & ARTCUST2 & " Set SELL_CODE = SELL_CODE_AC"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Sub Autofit_Notes()

        WorkbookView1.GetLock()
        workbook = WorkbookView1.ActiveWorkbook
        For i As Integer = 0 To workbook.Worksheets.Count - 1
            worksheet = workbook.Worksheets(i)
            If worksheet.Name <> "Summary" Then
                If worksheet.ProtectContents Then worksheet.Unprotect(XLS_PWD)

                Dim C As String = Excel_Cell0(-1, SCOL_Recap + 1)
                range = worksheet.Cells(C & ":" & C).EntireColumn
                If chkAutofitNotes.Checked Then
                    range.AutoFit()
                Else
                    range.ColumnWidth = 26.76
                End If

                For CBi2 As Integer = 1 To CBi

                    Dim C2 As String = Excel_Cell0(-1, SCOL_Recap + COLS_ST + CBi2 * COLS_Recap + 1)
                    range = worksheet.Cells(C2 & ":" & C2).EntireColumn
                    If chkAutofitNotes.Checked Then
                        range.AutoFit()
                    Else
                        range.ColumnWidth = 12
                    End If

                Next
                worksheet.Protect(XLS_PWD)
            End If
        Next
        WorkbookView1.ReleaseLock()

    End Sub

    Public Class MyCommandManager
        Inherits SpreadsheetGear.Commands.CommandManager
        Friend Sub New(workbookSet As SpreadsheetGear.IWorkbookSet)
            MyBase.New(workbookSet)
        End Sub
        Public Overrides Function CreateCommandPaste(range As SpreadsheetGear.IRange) As SpreadsheetGear.Commands.Command
            ' This is what would normally be called...
            ' return new CommandRange.Paste(range);  

            ' Anytime a Paste command is invoked, this will force a "Paste Values"
            Return New SpreadsheetGear.Commands.CommandRange.PasteSpecial(range, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
        End Function
    End Class

    'Private Sub cmdAutofitNotes_Click(sender As Object, e As EventArgs) Handles cmdAutofitNotes.Click
    '    Autofit_Notes(True)
    'End Sub


    Private Sub btnUnLockSalesBudget_Click(sender As Object, e As EventArgs) Handles btnUnLockSalesBudget.Click

        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
        Dim rowSPTMXWSS As DataRow = Fill_Record("SPTMXWSS", SEASON_CODE)
        If rowSPTMXWSS Is Nothing OrElse rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") & "" <> "1" Then
            MsgBox("Retail Sales Budgets are Not Locked", MsgBoxStyle.OkOnly, "Cannot Perform Request Action")
            Exit Sub
        Else
            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") = "0"
            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED_DATE") = DBNull.Value

            rowSPTMXWSS.Delete()
        End If

        If MsgBox("OK to Un-Lock Sales Budgets for " & SEASON_CODE & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        BeginTrans()

        INIT_LAST("SPTMXWSS", True, "", True)
        Update_Record_TDA("SPTMXWSS")

        TAC.TACMAIN1.Record_Event("SPTMXWSS", SEASON_CODE, Now, ASCMAIN1.USER_ID, "BUDUNL", "Sales Budgets are Un-Locked", SEASON_CODE, Me.Name)

        CommitTrans($"Sales Budgets for {SEASON_CODE} are now Un-Locked")

        Refresh_Documents()

    End Sub

    Private Sub btnLockSalesBudget_Click(sender As Object, e As EventArgs) Handles btnLockSalesBudget.Click

        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
        Dim rowSPTMXWSS As DataRow = Fill_Record("SPTMXWSS", SEASON_CODE)
        If rowSPTMXWSS Is Nothing Then
            rowSPTMXWSS = dst.Tables("SPTMXWSS").NewRow
            rowSPTMXWSS.Item("SEASON_CODE") = SEASON_CODE
            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") = "1"
            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED_DATE") = Now
            dst.Tables("SPTMXWSS").Rows.Add(rowSPTMXWSS)
        Else
            If rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") & "" = "1" Then
                MsgBox("Retail Sales Budgets are Already Locked", MsgBoxStyle.OkOnly, "Cannot Perform Request Action")
                Exit Sub
            End If

            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED") = "1"
            rowSPTMXWSS.Item("SEASON_SLS_BUD_LOCKED_DATE") = Now

        End If

        ASCMAIN1.sql = $"Select Count(*) RECS from SPTMXWS0 where SEASON_CODE = '{SEASON_CODE}'"
        Dim RECS As Integer = Val(ASCDATA1.GetDataValue)
        If RECS = 0 Then
            MsgBox($"No Weekly %'ages are on file for Season {SEASON_CODE}", MsgBoxStyle.OkOnly, "Cannot Lock Sales Budgets")
            Exit Sub
        End If

        If MsgBox("OK to Lock Sales Budgets for " & SEASON_CODE & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        BeginTrans()

        INIT_LAST("SPTMXWSS", True, "", True)
        Update_Record_TDA("SPTMXWSS")

        TAC.TACMAIN1.Record_Event("SPTMXWSS", SEASON_CODE, Now, ASCMAIN1.USER_ID, "BUDLCK", "Sales Budgets are Locked", SEASON_CODE, Me.Name)

        CommitTrans($"Sales Budgets for {SEASON_CODE} are now Locked")

        'MsgBox("Sales Budgets for " & SEASON_CODE & " are now Locked", MsgBoxStyle.OkOnly, "Verification")

        Refresh_Documents()

    End Sub

    Private Sub chkAutofitNotes_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutofitNotes.CheckedChanged
        Autofit_Notes()
    End Sub

    Private Sub WorkbookView1_MouseDown(sender As Object, e As MouseEventArgs) Handles WorkbookView1.MouseDown
        'If e.Button = Windows.Forms.MouseButtons.Left Then
        '    worksheet = WorkbookView1.ActiveSheet
        '    If worksheet.ProtectContents Then
        '        WorkbookView1.GetLock()
        '        worksheet.Unprotect(XLS_PWD)
        '        WorkbookView1.ReleaseLock()
        '    End If
        'End If
    End Sub

    Private Sub WorkbookView1_MouseUp(sender As Object, e As MouseEventArgs) Handles WorkbookView1.MouseUp
        'If e.Button = Windows.Forms.MouseButtons.Left Then
        '    worksheet = WorkbookView1.ActiveSheet
        '    If Not worksheet.ProtectContents Then
        '        WorkbookView1.GetLock()
        '        WorkbookView1.ActiveWorksheet.Protect(XLS_PWD)
        '        WorkbookView1.ReleaseLock()
        '    End If
        'End If
    End Sub

    Private Sub chkApproveAC_CheckedChanged(sender As Object, e As EventArgs) Handles chkApproveAC.CheckedChanged

    End Sub

    Private Sub grdSPTMXWSX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTMXWSX.InitializeRow
        If e.Row.Cells("DATA_UPDATED").Value & "" = "1" Then
            e.Row.Cells("SELL_CODE").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("SELL_CODE").ToolTipText = "Updates Pending Approval"
        End If
    End Sub

    Sub Initialize_FLWBs()

        Dim errors As New List(Of String)

        automate_Initialization = True

        ASCMAIN1.sql = "SELECT SELL_CODE, SELL_NAME FROM SOTSELL1 WHERE SELL_STATUS = 'A' AND SELL_CODE <> '000' ORDER BY SELL_CODE"
        For Each rowSOTSELL1 As DataRow In ASCDATA1.GetDataTable.Select("", "SELL_CODE")
            Dim SELL_CODE = rowSOTSELL1.Item("SELL_CODE")
            ASCMAIN1.Progress("Now Initializing", SELL_CODE)
            Absx1.txtFor("SELL_CODE").Text = SELL_CODE
            Click_Command("New")
            If Not ScreenMode Then
                errors.Add(SELL_CODE)
            Else
                Click_Command("Update")
            End If
        Next

        automate_Initialization = False

        If errors.Count <> 0 Then
            MsgBox("Some AEs could not be initialized: " & Join(errors.ToArray, "."), vbOKOnly, "Please Note")
        End If

        TAC.TACMAIN1.Record_Event("SPTMXWS1", SEASON_CODE, Now, ASCMAIN1.USER_ID, "SSNINI", $"Season Initialized for FLWBs", SEASON_CODE, Me.Name)

        ASCMAIN1.Progress("")

        MsgBox("Initialization Complete", MsgBoxStyle.OkOnly, "Success")

    End Sub
    Private Sub AutoFit_Summary_Columns()
        Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets("Summary")
        Dim firstCol As Integer = 4 ' Column E


        If workbook IsNot Nothing AndAlso workbook.WorkbookSet IsNot Nothing Then
            workbook.WorkbookSet.Calculate()
        End If

        Dim used = ws.UsedRange
        Dim lastCol As Integer = SCOL_Summary + (CBi * COLS_Summary) + (COLS_Summary - 1)

        ws.Range(0, firstCol, 0, firstCol).EntireColumn.ColumnWidth = 28

        For c As Integer = firstCol + 1 To lastCol

            Dim cell As SpreadsheetGear.IRange = ws.Cells(0, c)

            If cell.EntireColumn.Hidden Then Continue For

            If cell.ColumnWidth <= 3 Then Continue For

            cell.EntireColumn.AutoFit()

            cell.ColumnWidth = Math.Max(12, cell.ColumnWidth + 2)

        Next

    End Sub

End Class