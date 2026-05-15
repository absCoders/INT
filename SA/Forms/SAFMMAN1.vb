Public Class SAFMMAN1

    Dim RYP As String
    Dim FYP As String

    Dim BRAND_CODE As String

    Dim SATMMAN1 As String
    Dim SLS_MTD_TOTAL As Double
    Dim SLS_YTD_TOTAL As Double

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            ASCMAIN1.sql = "Select CODE1 CUST_CODE, CODE2 CUST_STORE_NO" _
                & ", SUM(FULL_2Y) FULL_2Y, SUM (FULL_LY) FULL_LY" _
                & ", SUM (YTD_TY) YTD_TY, SUM (YTD_LY) YTD_LY" _
                & " from WJZ_MKTMALL group by CODE1, CODE2"

            Create_TDA(.Tables.Add, "SATMMAN1", "**", 0, False, "", 0)

        End With

        grdSATMMAN1.DataSource = dst.Tables("SATMMAN1")

        Create_Summary(grdSATMMAN1, "CUST_CODE", "Count")

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP.ToString.Substring(0, 4) & "12"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")
                Validate_Code("BRAND_CODE")

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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Excel Extract"

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = tf
        'Setup_Summary()

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATMMAN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Reading from Sales History Data")
        Application.DoEvents()

        Save_Header_Fields(UltraGroupBox1)

        Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
        RYP = z
        FYP = ASCMAIN1.Period_Calc(z, -11)

        dst.EnforceConstraints = False

        Fill_Records("SATMMAN1")
        'EnforceConstraints(True)

        ASCMAIN1.Progress("Now Setting Up Screen")
        Setup_tabMain()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATMMAN1, "SSB", "Show Filter", "Show GroupBox", "Customer Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        Select Case e.SourceControl.Name
            Case "grdSATSLSC1"

            Case Else
        End Select
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
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Absx1.txtFor("BRAND_CODE").Text <> "" Then
                        Click_Command("Load", e)
                    End If
                End If
            Case "BRAND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Absx1.txtFor("OPS_YYYYPP").Text <> "" Then
                        Click_Command("Load", e)
                    End If
                End If
        End Select
    End Sub
#End Region

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()
        Print_Report_End()
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
        'If tabMain.SelectedTab Is Nothing Then
        '    UltraExplorerBar1.Groups("Summaries").Visible = False
        'Else
        '    UltraExplorerBar1.Groups("Summaries").Visible = (tabMain.SelectedTab.Key = "Summaries")
        '    Setup_Summary()
        'End If
    End Sub


    Sub Create_Allocation_Status_Tables(Optional initialize As Boolean = False)

        If initialize Then
        Else
            ASCMAIN1.Progress("Now Creating Allocations Status Tables")
        End If

        Dim sql_ictitem1 As String
        Dim DT As String = "" 'Format(dteAllocations.Value, "dd-MMM-yyyy")

        ASCMAIN1.sql = "Select SOTALLO1.*" & vbCrLf _
            & ", '000000' TYP_START, '000000' TYP_END" & vbCrLf _
            & ", '000000' LYP_START, '000000' LYP_END" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & sql_ICTITEM1 & vbCrLf _
            & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_PICK" & vbCrLf _
            & ", 0 QTY_ALLO_TOTAL" & vbCrLf _
            & " from SOTALLO1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTALLO1.DATE_START <= '" & DT & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END   >= '" & DT & "'"


        If initialize Then
            'ASCMAIN1.sql &= " and ROWNUM < 1"
            'SOTALLO1S = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO1S & " Add Primary Key (ALLO_CTL_NO)")
        Else
            'ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO1S)
            'ASCDATA1.ExecuteSQL("Insert into " & SOTALLO1S & " " & ASCMAIN1.sql)


            'ASCMAIN1.sql = "Update " & SOTALLO1S & " Set TYP_START = (Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > DATE_START)"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update " & SOTALLO1S & " Set TYP_END = (Select Min (OPS_YYYYPP) from GLTPARM2 where PRD_END_DATE > DATE_END)"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update " & SOTALLO1S & " Set LYP_START = PERIOD_CALC(TYP_START,-12)"
            'ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update " & SOTALLO1S & " Set LYP_END = PERIOD_CALC(TYP_END,-12)"
            'ASCDATA1.ExecuteSQL()



            'ASCMAIN1.sql = "" _
            '    & "Begin" & vbCrLf _
            '    & " Declare Cursor C1 is" & vbCrLf _
            '    & "  Select ITEM_CODE" & vbCrLf _
            '    & "  , Sum (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
            '    & "  , Sum (WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
            '    & "   from ICTSTAT2" & vbCrLf _
            '    & "   where ITEM_CODE in (Select ITEM_CODE from " & SOTALLO1S & ")" & vbCrLf _
            '    & "  group by ITEM_CODE;" & vbCrLf _
            '    & " Begin" & vbCrLf _
            '    & "  For R1 in C1 Loop" & vbCrLf _
            '    & "   Update " & SOTALLO1S & " Set " & vbCrLf _
            '    & "      WHSE_QTY_ON_HAND = R1.WHSE_QTY_ON_HAND" & vbCrLf _
            '    & "     ,WHSE_QTY_PICK = R1.WHSE_QTY_PICK" & vbCrLf _
            '    & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            '    & "  End Loop;" & vbCrLf _
            '    & " End;" & vbCrLf _
            '    & "End;"
            ASCDATA1.ExecuteSQL()
        End If


        'ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
        '    & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
        '    & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
        '    & " from " & SOTALLO1S & " SOTALLO1,SOTALLO2,ARTCUST1" & vbCrLf _
        '    & " where SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
        '    & "   and ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE"
        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            'SOTALLO2S = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTALLO2S & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")
        Else
            'ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLO2S)
            'ASCDATA1.ExecuteSQL("Insert into " & SOTALLO2S & " " & ASCMAIN1.sql)
        End If



        'ASCMAIN1.sql = "Select SOTALLO3.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
        '    & ", 0 ORDR_QTY, 0 ORDR_QTY_OPEN, 0 ORDR_QTY_PICK, 0 ORDR_QTY_SHIP, 0 ORDR_QTY_CANC" & vbCrLf _
        '    & ", 0 LY_QTY_SELL_IN1, 0 LY_QTY_SELL_IN2, 0 LY_QTY_SELL_THRU1, 0 LY_QTY_SELL_THRU2" & vbCrLf _
        '    & " from " & SOTALLO1S & " SOTALLO1,SOTALLO3,ARTCUST1,ARTCUST2" & vbCrLf _
        '    & " where SOTALLO3.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO" & vbCrLf _
        '    & "   and ARTCUST1.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUST2.CUST_CODE = SOTALLO3.CUST_CODE" & vbCrLf _
        '    & "   and ARTCUST2.CUST_STORE_NO = SOTALLO3.CUST_STORE_NO"
        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
            'SOTALLO3S = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & "SOTALLO3S" & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & "SOTALLO3S")
            ASCDATA1.ExecuteSQL("Insert into " & "SOTALLO3S" & " " & ASCMAIN1.sql)

        End If

        If Not initialize Then
            ' Get_Sales_STATS(True)

            'ASCMAIN1.sql = "Update " & "SOTALLO1S" & " SOTALLO1S Set QTY_ALLO_TOTAL = (Select Sum (QTY_ALLO) from " & SOTALLO2S & " SOTALLO2S where SOTALLO2S.ALLO_CTL_NO = SOTALLO1S.ALLO_CTL_NO)"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.Progress("")

    End Sub
    Sub Generate_Pivot()

        Dim useSSG As Boolean = False ' chkUseSSG.Checked

        Create_Allocation_Status_Tables()

        ASCMAIN1.Progress("Now Creating Allocations Status Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If useSSG Then
            workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            worksheet = workbook.Sheets("VLOOKUP")
        Else
            excel = New Microsoft.Office.Interop.Excel.Application
            wb = excel.Workbooks.Open(FILENAME)
            ws = wb.Worksheets("VLOOKUP")
        End If


        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & " from ICTCOLL1,ICTBRAN1 " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " order by COLLECTION_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("K4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Brand_List", "=VLOOKUP!$K$3:$O$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "COLLECTION_CODE")
                r += 1
                ws.Range("K" & CStr(3 + r) & ":O" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Brand_List", "=VLOOKUP!$K$3:$O$" & CStr(3 + DataTable.Rows.Count))
        End If

        Dim DT As String
        'DT = Format(dteAllocations.DateTime.Date, "dd-MMM-yyyy")
        ' DT = dteAllocations.DateTime.ToString("dd-MMM-yyyy")

        'D4=VLOOKUP(C4,$K$4:$L$97,2,FALSE)
        'D4=VLOOKUP(C4,Brand_List,2,FALSE)
        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, '=VLOOKUP(C' || TRIM(TO_CHAR(3 + ROWNUM)) || ',Brand_List,2,FALSE)' COLLECTION_NAME" & vbCrLf _
            & ", CASE WHEN NVL(ICTITEM1.COST_CATGY_CODE,'?') = 'S' THEN 'RETAIL' ELSE 'COLLATERAL' END RC_TYPE" & vbCrLf _
            & ", ICTITEM1.ITEM_DATE_TO_SHIP" & vbCrLf _
            & ", ICTITEM1.PROD_CODE, 'H1' HALF, A.QTY_ALLO" & vbCrLf _
            & " from ICTITEM1, " & vbCrLf _
            & "(Select SOTALLO1.ITEM_CODE, Sum (SOTALLO2.QTY_ALLO) QTY_ALLO" & vbCrLf _
            & " from SOTALLO1,SOTALLO2" & vbCrLf _
            & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
            & "   and SOTALLO1.DATE_START <= '" & DT & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END   >= '" & DT & "'" & vbCrLf _
            & "   group by SOTALLO1.ITEM_CODE" & vbCrLf _
            & "   order by SOTALLO1.ITEM_CODE) A" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = A.ITEM_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("A4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Item_List", "=VLOOKUP!$A$3:$I$" & CStr(3 + DataTable.Rows.Count))
            worksheet.Cells("D4:D4").Copy(worksheet.Cells("D5:D" & CStr(3 + DataTable.Rows.Count)))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "ITEM_CODE")
                r += 1
                ws.Range("A" & CStr(3 + r) & ":I" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Item_List", "=VLOOKUP!$A$3:$I$" & CStr(3 + DataTable.Rows.Count))

            xlSourceRange = ws.Range("D4:D4")
            xlDestRange = ws.Range("D4:D" & CStr(3 + DataTable.Rows.Count))
            xlSourceRange.Copy(xlDestRange)
        End If





        Dim CUST_CODE_AE As String = "IPLBAE" ' Probably should be Parameterized

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from (" & vbCrLf _
            & "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & " where CUST_CODE IN (Select Distinct CUST_CODE from " & "SOTALLO2S" & ")" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO CUST_CODE," & vbCrLf _
            & " ARTCUST2.CUST_STORE_NAME, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from ARTCUST1,ARTCUST2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = '" & CUST_CODE_AE & "'" & vbCrLf _
            & ") order by CUST_CODE"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("Q4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("Chain_List", "=VLOOKUP!$Q$3:$S$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("", "CUST_CODE")
                r += 1
                ws.Range("Q" & CStr(3 + r) & ":S" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Chain_List", "=VLOOKUP!$Q$3:$S$" & CStr(3 + DataTable.Rows.Count))
        End If


        If useSSG Then
            worksheet.Visible = False
        Else
            ws.Visible = False
        End If

        If useSSG Then
            worksheet = workbook.Sheets("DATA")
        Else
            ws = wb.Worksheets("DATA")
        End If

        'ASCMAIN1.sql = "Select ARTCUST1.CUST_NAME, SOTALLO2.CUST_CODE" & vbCrLf _
        '    & ", TO_CHAR(SOTALLO1.DATE_START,'MMDDYY') DATE_START" & vbCrLf _
        '    & ", SOTALLO1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
        '    & ", ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_RETAIL_PRICE, SOTALLO2.QTY_ALLO" & vbCrLf _
        '    & ", SOTALLO2.ORDR_QTY_SHIP, SOTALLO2.ORDR_QTY_OPEN, ICTITEM1.ITEM_SO_QTY_MULT" & vbCrLf _
        '    & ", SOTALLO1.WHSE_QTY_ON_HAND, SOTALLO1.WHSE_QTY_PICK" & vbCrLf _
        '    & " from " & SOTALLO1S & " SOTALLO1," & SOTALLO2S & " SOTALLO2,ARTCUST1,ICTITEM1" & vbCrLf _
        '    & "where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE " & vbCrLf _
        '    & "  and ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
        '    & "  and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO"
        DataTable = ASCDATA1.GetDataTable

        If useSSG Then
            range = worksheet.Cells("R4")
            range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
            workbook.Names.Add("PivotBase", "=DATA!$R$3:$AD$" & CStr(3 + DataTable.Rows.Count))
        Else
            r = 0
            For Each row As DataRow In DataTable.Select("")
                r += 1
                ws.Range("R" & CStr(3 + r) & ":AD" & CStr(3 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("PivotBase", "=DATA!$R$3:$AD$" & CStr(3 + DataTable.Rows.Count))
        End If

        If useSSG Then
            worksheet.Cells("A4:Q4").Copy(worksheet.Cells("A4:Q" & CStr(3 + DataTable.Rows.Count)))
            worksheet.Cells("C1").Value = Now
        Else
            xlSourceRange = ws.Range("A4:Q4")
            xlDestRange = ws.Range("A4:Q" & CStr(3 + DataTable.Rows.Count))
            xlSourceRange.Copy(xlDestRange)
            ws.Cells(1, 3).Value = Now
        End If

        If useSSG Then
            worksheet.Visible = False
        Else
            ws.Visible = False
        End If

        If useSSG Then

        Else
            excel.Run("ResetData")
        End If


        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Allocations"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"

                If useSSG Then
                    workbook.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbookMacroEnabled) ' SpreadsheetGear.FileFormat.OpenXMLWorkbook) ' SpreadsheetGear.FileFormat.OpenXMLWorkbookMacroEnabled)
                Else
                    Dim objOpt As Object = Nothing ' Missing.Value
                    wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                              , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                    wb.Close(False, objOpt, objOpt)
                End If

                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        If useSSG Then
        Else
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
        End If

        'Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)
        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")
    End Sub
End Class