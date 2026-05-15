Public Class RSFSSPC1

#Region "General Declarations"
    Dim rowRSTSSPC1 As DataRow

    Dim rowARTCUST1 As DataRow
    Dim rowICTSEAS1 As DataRow
    Dim rowSOTTCLS1 As DataRow
    Dim rowSOTPCLS1 As DataRow

    Dim CUST_CODE As String
    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String

    Dim HC_CODE As String

    Dim SEASON_CODE_prior As String
    Dim SEASON_CODE_LY As String
    ' Dim SEASON_is_current_or_future As Boolean

    Dim sqlSOTALLOX As String

    Dim sLINE_TAG As Integer = -1
    Dim eLINE_TAG As Integer = -1

    Dim XLSR As New Dictionary(Of String, Integer)
    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim xls_COLLECTION_CODEs As New List(Of String)

    Dim Delete_Sheets As New List(Of String)

    Dim DTE1 As Date
    Dim DTE2 As Date
    Dim XLS_NO As String
    Dim XLS_PWD As String = "ABS"
    Dim XLS_Allocation_Lines As New Dictionary(Of String, Integer)
    Dim YPs_Imported As New List(Of String)

    Dim RSTSSPC2 As String

    ' LIST OF LINE TAGS DEFINING THE LINES THAT CONTAIN 7 MOS OF NUMERIC DATA THAT SHOULD BE SAVED WITH SSP
    Dim LINE_TAGs As List(Of String) = _
        {"TYPBOM", "TYPSLS", "TYPSLSACT", "TYPSLSACTZ", _
         "TYPGRSB", "TYPGRSP", "TYPGRS", "TYPADJ", "TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET", "TYPEOM", _
         "TYPSHP%LY", "TYPSLS%LY", _
         "CUSEOM", "CUSRTL", "CUSSTC", _
         "TYOBOM", "TYOSLS", "TYOGRS", "TYOADJ", "TYOEOM", _
         "LYABOM", "LYASLS", "LYAGRS", "LYAADJ", "LYAEOM", _
         "TYPB1", "TYPB2", "TYPB3", "TYPB4", "TYPB5", "TYPB6", "TYPB7", "TYPB8", "TYPB9", _
         "TYAGRSB", "TYAGRSP", "TYAADJ", _
         "TYPGRSBT", "TYPGRSPT", "TYPADJT", _
         "TYPDAM", "TYPDIS", "TYPOVR", "TYPSET", "TYPCLS", "PIPE", "CARRYOVR"}.ToList

    Dim YPs() As String
    Dim c0 As Integer = 4 ' OFFSET FOR 1ST MONTH OF DATA, COL E = 4

    Dim CUST_CODEs_import As String = ""
    Dim ITEM_CODEs_import As String = ""
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)
            Create_TDA(.Tables.Add, "RSTSSPC1", "*", 3)
            Create_TDA(.Tables.Add, "RSTSSPC2", "*", 3)

            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTCOLL1", "*", 0, False)

            ASCMAIN1.sql = "Select SATBUDW1.*,SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
                & " from SATBUDW1,ICTCOLL1,ARTCUST1,SOTTCLS1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = SATBUDW1.COLLECTION_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SATBUDW1.CUST_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE"
            Create_TDA(.Tables.Add, "SATBUDW1", "*", 0)
            Dim TX As String = ""
            For I As Integer = 1 To 12
                Dim C As String = "WB_P" & Format(I, "00")
                .Tables("SATBUDW1").Columns(C).DefaultValue = 0
                TX &= "+ISNULL(" & C & ",0)"
            Next
            .Tables("SATBUDW1").Columns.Add("WB_P00", GetType(System.Decimal), Mid(TX, 2))

            ASCMAIN1.sql = "Select RSTSSPC1.*,SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                & " from RSTSSPC1,ARTCUST1,SOTTCLS1" & vbCrLf _
                & " where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = RSTSSPC1.CUST_CODE" & vbCrLf _
                & "   and SEASON_CODE = :PARM1"
            Create_TDA(.Tables.Add, "RSTSSPCX", "**", 0, False, "V", 3)
            .Tables("RSTSSPCX").Columns.Add("SEL")
            .Tables("RSTSSPCX").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTALLO1.*, SOTALLO2.QTY_ALLO" _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_RETAIL_PRICE" _
                & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE" _
                & ", ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" _
                & " from SOTALLO1,SOTALLO2,ICTITEM1,ICTCOLL1" _
                & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
                & "   and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO"
            sqlSOTALLOX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False)

            With .Tables.Add("SATBUDWX")
                .Columns.Add("CHANNEL_CODE")
                .Columns.Add("COLLECTION_CODE")
                For Each PFX As String In New String() {"R", "P", "S", "F", "X"}
                    Dim T As String = ""
                    For I As Integer = 0 To 6
                        Dim C As String = "BUD_" & Format(I, "0")
                        .Columns.Add(PFX & C, GetType(System.Decimal))
                        If I <> 0 Then T &= "+ISNULL(" & PFX & C & ",0)"
                        If PFX = "X" Then
                            .Columns(PFX & C).Expression = "ISNULL(F" & C & ",0) - ISNULL(S" & C & ",0)"
                        End If
                    Next
                    .Columns(PFX & "BUD_0").Expression = Mid(T, 2)
                Next
                .PrimaryKey = New DataColumn() {.Columns("CHANNEL_CODE"), .Columns("COLLECTION_CODE")}
            End With

            ' NOTE - THE LINE TAGS IN THE WHERE CLAUSE MAY NEED TO BE EXPANDED IF WE EVER ADD LINES TO THE ROLLUP
            ASCMAIN1.sql = "Select COLLECTION_CODE, LINE_TAG" & vbCrLf _
                & ", SUM (NVL(AMT_0,0)+NVL(AMT_1,0)+NVL(AMT_2,0)+NVL(AMT_3,0)+NVL(AMT_4,0)+NVL(AMT_5,0)) AMT" & vbCrLf _
                & " from RSTSSPC2 " & vbCrLf _
                & " where SEASON_CODE = '" & SEASON_CODE_prior & "'" & vbCrLf _
                & "   and HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   and LINE_TAG IN ('TYPSLSACTZ','TYPGRSB','TYPGRSP','TYPGRS','TYPGRSNET','TYPEOM')" & vbCrLf _
                & " group by COLLECTION_CODE, LINE_TAG"
            Create_TDA(.Tables.Add, "RSTSSPCR", "**", 0, False, , 2)
        End With

        Fill_Records("ICTCOLL0")
        Fill_Records("ICTCOLL1")

        grdRSTSSPCX.DataSource = dst.Tables("RSTSSPCX")
        Create_Summary(grdRSTSSPCX, "CUST_CODE", "Count")
        Create_Summary(grdRSTSSPCX, "SEL")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTSSPCX.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                gcol.Hidden = True ' NO CONSOLIDATIONS
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        grdRSTSSPCX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        Dim YYYY As String = Mid(ASCMAIN1.CYP, 1, 4)
        Dim NY As String = Format(Val(YYYY) + 1, "0000")
        Dim LY As String = Format(Val(YYYY) - 1, "0000")

        ASCMAIN1.sql = "Select Min (SEASON_CODE) from RSTSSPC1"
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
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("CUST_CODE")
                Validate_Code("HC_CODE")

                If EMsg = "" Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    Dim HC_CODE As String = Absx1.txtFor("HC_CODE").Text
                    Dim row As DataRow = LookUp("RSTSSPC1", New String() {CUST_CODE, SEASON_CODE, HC_CODE})
                    If row IsNot Nothing Then
                        EMsg &= vbCr & "Record Already Exists for Customer " & CUST_CODE & " in Season " & SEASON_CODE & " for HC " & HC_CODE
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("RSTSSPC1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "View"
                Validate_Code("CUST_CODE")
                Validate_Code("HC_CODE")

                If EMsg = "" Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    Dim HC_CODE As String = Absx1.txtFor("HC_CODE").Text
                    Dim row As DataRow = LookUp("RSTSSPC1", New String() {CUST_CODE, SEASON_CODE, HC_CODE})
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record on file for Customer " & CUST_CODE & " in Season " & SEASON_CODE & " for HC " & HC_CODE
                    End If
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("RSTSSPC1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Update"
                WorkbookView1.EndEdit()
                Process_Workbook()

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

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Save XLSX"
                WorkbookView1.GetLock()
                Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("RSFSSPC1.XLSX_NO") & ".XLSX"
                WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                Show_Document(FILENAME)
                WorkbookView1.ReleaseLock()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Save XLSX").Visible = ScreenMode

                    If EntryMode = "V" And ScreenMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L")

                    .Items("New").Visible = Not .Items("Update").Visible
                    .Items("Edit").Visible = Not .Items("Update").Visible

                    .Items("Done").Visible = ScreenMode And EntryMode = "V"
                End With

                .Groups("Display Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdRSTSSPCX.Visible = Not ScreenMode
        splSSPC.Visible = ScreenMode

        '  spl.Panel1Collapsed = (EntryMode = "P")

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "RSTSSPC1", "RSTSSPC2", "SATBUDWX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        SEASON_CODE = ""
        HC_CODE = ""
        chkShowLINE_TAGs.Checked = False
         
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("HC_CODE").Text = ""

        If Absx1.cbeFor("SEASON_CODE").Value & "" = "" Then
            Dim YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)
            Absx1.cbeFor("SEASON_CODE").Value = Mid(YP, 1, 4) & IIf(Mid(YP, 5, 2) < "07", "S", "F")
        End If

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        SEASON_CODE = HFs("SEASON_CODE")
        HC_CODE = HFs("HC_CODE")

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        SEASON_TYPE = rowICTSEAS1.Item("SEASON_TYPE")
        SEASON_YEAR = rowICTSEAS1.Item("SEASON_YEAR")
        DTE1 = CDate(IIf(SEASON_TYPE = "S", "02", "08") & "/01/" & SEASON_YEAR)
        DTE2 = DTE1.AddMonths(6).AddDays(-1)

        SEASON_CODE_prior = IIf(SEASON_TYPE = "S", _
                                Format(Val(SEASON_YEAR) - 1, "0000") & "F", _
                                SEASON_YEAR & "S")
        SEASON_CODE_LY = Format(Val(SEASON_YEAR) - 1, "0000") & SEASON_TYPE

        ReDim YPs(6) ' 0 is either Jan or Jul
        If SEASON_TYPE = "S" Then
            YPs(0) = SEASON_YEAR & "01"
        Else
            YPs(0) = SEASON_YEAR & "07"
        End If
        For I As Integer = 1 To 6
            YPs(I) = ASCMAIN1.Period_Calc(YPs(0), I)
        Next

        If EntryMode = "N" Then
            rowRSTSSPC1 = dst.Tables("RSTSSPC1").NewRow
            rowRSTSSPC1.Item("CUST_CODE") = CUST_CODE
            rowRSTSSPC1.Item("SEASON_CODE") = SEASON_CODE
            rowRSTSSPC1.Item("HC_CODE") = HC_CODE
            rowRSTSSPC1.Item("INIT_DATE") = DATETIME_STAMP
            rowRSTSSPC1.Item("INIT_OPER") = ASCMAIN1.USER_ID

            XLS_NO = ASCMAIN1.Next_Control_No("RSTSSPC1.XLS_NO")
            rowRSTSSPC1.Item("XLS_NO") = XLS_NO

            dst.Tables("RSTSSPC1").Rows.Add(rowRSTSSPC1)
        Else
            rowRSTSSPC1 = Fill_Record("RSTSSPC1", New String() {CUST_CODE, SEASON_CODE, HC_CODE})
            XLS_NO = rowRSTSSPC1.Item("XLS_NO") & ""
        End If

        EnforceConstraints(False)

        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        rowSOTTCLS1 = LookUp("SOTTCLS1", rowARTCUST1.Item("TRADE_CLASS_CODE"))
        rowSOTPCLS1 = LookUp("SOTPCLS1", rowARTCUST1.Item("PRICE_CLASS_CODE"))

        EnforceConstraints(True)

        If EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V" Then
            Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
            WorkbookView1.GetLock()
            WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            XLS_Validation(True)

            XLS_Refresh_COLLECTION()
            XLS_Refresh_Allocations()

            Set_Month_Headings(SEASON_CODE)

            For Each SHEET_NAME As String In Delete_Sheets
                workbook.Worksheets(SHEET_NAME).Delete()
            Next

            For i As Integer = 0 To workbook.Worksheets.Count - 1
                With workbook.Worksheets(i)
                    .Range(0, 0, 0, 2).EntireColumn.Hidden = True
                End With
            Next

            If EntryMode = "E" Or EntryMode = "V" Then

                Fill_Records("RSTSSPC2", New String() {CUST_CODE, SEASON_CODE, HC_CODE})

                Dim COLLECTION_CODE As String = ""
                For Each row As DataRow In dst.Tables("RSTSSPC2").Select("", "COLLECTION_CODE")
                    If row.Item("COLLECTION_CODE") <> COLLECTION_CODE Then
                        COLLECTION_CODE = row.Item("COLLECTION_CODE")
                        worksheet = workbook.Worksheets(COLLECTION_CODE)
                    End If

                    If worksheet IsNot Nothing Then
                        Dim LINE_TAG As String = row.Item("LINE_TAG")
                        '  If LINE_TAG = "TYPGRSBNET" And ASCMAIN1.Running_in_VS Then Stop
                        If LINE_TAG = "PSIITM" Or LINE_TAG = "TOAITM" Then
                            Dim LINE_KEY As String = row.Item("LINE_KEY")
                            If XLS_Allocation_Lines.Keys.Contains(LINE_KEY) Then
                                Dim r As Integer = XLS_Allocation_Lines(LINE_KEY)
                                For m As Integer = 0 To 6
                                    If worksheet.Cells(r, c0 + m).Formula & "" = "" Then
                                        worksheet.Cells(r, c0 + m).Value = row.Item("AMT_" & CStr(m))
                                    End If
                                Next
                            End If

                        Else

                            If LINE_TAG = "TYPSLS%LY" Or LINE_TAG = "TYPSHP%LY" Then
                            Else
                                For m As Integer = 0 To 6
                                    If worksheet.Cells(XLSR(LINE_TAG), c0 + m).Formula & "" = "" Then
                                        worksheet.Cells(XLSR(LINE_TAG), c0 + m).Value = row.Item("AMT_" & CStr(m))
                                    End If
                                Next
                            End If

                            If New String() {"TYPBOM", "LYABOM"}.Contains(LINE_TAG) Then
                                worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + 3).Value = row.Item("AMT_X")
                            End If

                            If New String() {"TYPBOM", "TYPSLS", "TYPSLSACT", "TYPSLSACTZ", "TYPGRSB", "TYPGRSP", "TYPGRS", "TYPADJ", "TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET", "TYPEOM", "TYPSHP%LY", "TYPSLS%LY"}.Contains(LINE_TAG) Then
                                worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + 3 + 4).Value = row.Item("NOTES")
                            End If

                        End If
                    End If
                Next

                If EntryMode = "V" Then
                    ' Stop ' WHY IS CUST_CODE ABSENT?
                    ASCMAIN1.sql = "Select COLLECTION_CODE, LINE_TAG" & vbCrLf _
                        & ", SUM (NVL(AMT_0,0)+NVL(AMT_1,0)+NVL(AMT_2,0)+NVL(AMT_3,0)+NVL(AMT_4,0)+NVL(AMT_5,0)) AMT" & vbCrLf _
                        & " from RSTSSPC2 " & vbCrLf _
                        & " where SEASON_CODE = '" & SEASON_CODE_prior & "'" & vbCrLf _
                        & "   and HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                        & "   and CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & "   and LINE_TAG IN ('TYPSLSACTZ','TYPGRSB','TYPGRSP','TYPGRS','TYPGRSNET','TYPEOM')" & vbCrLf _
                        & " group by COLLECTION_CODE, LINE_TAG"
                    Fill_Records("RSTSSPCR", "", True, ASCMAIN1.sql)
                End If
            End If

            WorkbookView1.ReleaseLock()

            'If EntryMode = "N" Then
            Get_LY_Actuals()
            'End If
        Else
            WorkbookView1.GetLock()
            XLS_Validation(False)
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


        Get_TY_Actuals()

        WorkbookView1.GetLock()
        For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
            If worksheet.ProtectContents Then worksheet.Unprotect(XLS_PWD)
            worksheet.Cells(XLSR("TYPGRSBNET"), 0).EntireRow.Hidden = True
            worksheet.Cells(XLSR("TYPGRSPNET"), 0).EntireRow.Hidden = True

            worksheet.Protect(XLS_PWD)
        Next
        WorkbookView1.ReleaseLock()


        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Dim sqld As String = "CUST_CODE = '" & CUST_CODE & "' and SEASON_CODE = '" & SEASON_CODE & "' and HC_CODE = '" & HC_CODE & "'"
        INIT_LAST("RSTSSPC1")
        Update_Record_TDA("RSTSSPC1", sqld)
        Update_Record_TDA("RSTSSPC2", sqld)
        CommitTrans("Update Complete")

    End Sub

    Sub Process_Workbook()

        dst.Tables("RSTSSPC2").Rows.Clear()

        WorkbookView1.GetLock()

        For i As Integer = 1 To workbook.Worksheets.Count - 1
            With workbook.Worksheets(i)
                If Delete_Sheets.Contains(.Name) Then
                Else
                    Try
                        For Each LINE_TAG As String In LINE_TAGs
                            Save_LINE_TAG(LINE_TAG, workbook.Worksheets(i))
                        Next

                        Get_Range("PSIITM")
                        For r As Integer = sLINE_TAG To eLINE_TAG
                            If .Cells(r, 2).Value & "" = .Cells(1, 1).Value & "" Then
                                Save_LINE_TAG("PSIITM", workbook.Worksheets(i), r)
                            End If
                        Next

                        Get_Range("TOAITM")
                        For r As Integer = sLINE_TAG To eLINE_TAG
                            If .Cells(r, 2).Value & "" = .Cells(1, 1).Value & "" Then
                                Save_LINE_TAG("TOAITM", workbook.Worksheets(i), r)
                            End If
                        Next
                    Catch ex As Exception
                        EMsg &= "Error occured in Sheet " & workbook.Worksheets(i).Name & vbCrLf & ex.Message
                    End Try
                End If
            End With
        Next

        WorkbookView1.ReleaseLock()
    End Sub

    Function Get_SQL_SATBUDW1() As String
        Dim sqld As String = "SATBUDW1.OPS_YYYY = '" & SEASON_YEAR & "'"
        Return sqld
    End Function
    Sub Save_LINE_TAG(LINE_TAG As String, _
                    ws As SpreadsheetGear.IWorksheet, _
                    Optional XLR As Integer = -1)

        Dim LINE_KEY As String = "X"
        If XLR = -1 Then
            XLR = XLSR(LINE_TAG)
        Else
            LINE_KEY = ws.Cells(XLR, 1).Value
        End If

        Dim row As DataRow = dst.Tables("RSTSSPC2").NewRow
        row.Item("CUST_CODE") = CUST_CODE
        row.Item("SEASON_CODE") = SEASON_CODE
        row.Item("HC_CODE") = HC_CODE
        row.Item("COLLECTION_CODE") = ws.Name
        row.Item("LINE_TAG") = LINE_TAG
        row.Item("LINE_KEY") = LINE_KEY

        Dim something_to_save As Boolean = False

        If LINE_TAG = "TYPSLS%LY" Or LINE_TAG = "TYPSHP%LY" Then
        Else
            For m As Integer = 0 To 6
                If ws.Cells(XLR, c0 + m).Value & "" <> "" Then
                    Dim AMT As Decimal = Val(ws.Cells(XLR, c0 + m).Value)
                    row.Item("AMT_" & CStr(m)) = ws.Cells(XLR, c0 + m).Value
                    something_to_save = True
                End If
            Next
        End If

        If New String() {"TYPBOM", "LYABOM"}.Contains(LINE_TAG) Then
            If ws.Cells(XLR, c0 + 6 + 3).Value & "" <> "" Then
                Dim AMT As Decimal = Val(ws.Cells(XLR, c0 + 6 + 3).Value)
                row.Item("AMT_X") = ws.Cells(XLR, c0 + 6 + 3).Value
                something_to_save = True
            End If
        End If

        If New String() {"TYPBOM", "TYPSLS", "TYPSLSACT", "TYPSLSACTZ", "TYPGRSB", "TYPGRSP", "TYPGRS", "TYPADJ", "TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET", "TYPEOM", "TYPSHP%LY", "TYPSLS%LY"}.Contains(LINE_TAG) Then
            If ws.Cells(XLR, c0 + 6 + 3 + 4).Value & "" <> "" Then
                row.Item("NOTES") = ws.Cells(XLR, c0 + 6 + 3 + 4).Value
                something_to_save = True
            End If
        End If

        dst.Tables("RSTSSPC2").Rows.Add(row)
        If Not something_to_save Then
            row.Delete()
        End If

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
        Load_Popup_Menu(grdRSTSSPCX, "SS", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Select All for Channel")
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
                Case "grdRSTSSPCX"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                    tlb_btn = DirectCast(tlb_pop.Tools("Select All for Channel"), UltraWinToolbars.ButtonTool)
                    If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim CHANNEL_CODE As String = grd.ActiveRow.Cells("CHANNEL_CODE").Value & ""
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Select All for Channel " & CHANNEL_CODE
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdRSTSSPCX.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All for Channel"
                Dim CHANNEL_CODE As String = grd.ActiveRow.Cells("CHANNEL_CODE").Value & ""
                Dim sqlw As String = "CHANNEL_CODE = '" & CHANNEL_CODE & "'"
                For Each row As DataRow In dst.Tables("RSTSSPCX").Select(sqlw)
                    row.Item("SEL") = "1"
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
            Case "HC_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
            Case "COLLECTION_CODE"
                If ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("COLLECTION_CODE").Text <> "" Then
                        If LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text) IsNot Nothing Then
                            XLS_Refresh_COLLECTION(Absx1.txtFor("COLLECTION_CODE").Text)
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
            Case "COLLECTION_CODE"
                If Absx1.txtFor("COLLECTION_CODE").Text <> "" Then
                    If LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text) IsNot Nothing Then
                        XLS_Refresh_COLLECTION(Absx1.txtFor("COLLECTION_CODE").Text)
                    End If
                End If
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
            Case "HC_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("HC_CODE").Text <> "" Then
                        LookUp("ICTCOLL0", Absx1.txtFor("HC_CODE").Text)
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

    Sub Set_Month_Headings(SEASON_CODE As String)
        Dim SD As Date = CDate(IIf(SEASON_TYPE = "S", "02", "08") & "/01/" & SEASON_YEAR)

        Dim rMOSHDG As Integer = XLSR("MOSHDG")
        For M As Integer = 0 To 6
            Dim D As Date = SD.AddMonths(M - 1)
            Dim LEGEND As String = Format(D, "MMM")
            workbook.Worksheets(0).Cells(rMOSHDG, c0 + M).Value = LEGEND
        Next

        Dim PRICE_BASE_DPCT As Decimal = 0

        Dim rCUST_CODE As Integer = XLSR("CUST_CODE")
        workbook.Worksheets(0).Cells(rCUST_CODE, 1).Value = CUST_CODE
        workbook.Worksheets(0).Cells(rCUST_CODE, 2).Value = Absx1.txtFor("CUST_NAME").Text
        PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
        workbook.Worksheets(0).Cells(XLSR("TYPGRSNET"), 1).Value = 100 - PRICE_BASE_DPCT

        Dim rSEASON_CODE As Integer = XLSR("SEASON_CODE")
        workbook.Worksheets(0).Cells(rSEASON_CODE, 1).Value = SEASON_CODE
        workbook.Worksheets(0).Cells(rSEASON_CODE, 2).Value = rowICTSEAS1.Item("SEASON_DESC")

        Dim rowICTSEAS1_LY As DataRow = LookUp("ICTSEAS1", SEASON_CODE_LY)
        Dim rLYAHDG As Integer = XLSR("LYAHDG")
        workbook.Worksheets(0).Cells(rLYAHDG, 1).Value = rowICTSEAS1_LY.Item("SEASON_CODE")
        workbook.Worksheets(0).Cells(rLYAHDG, 2).Value = rowICTSEAS1_LY.Item("SEASON_DESC")

    End Sub

    Sub Refresh_Documents()
        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
        Fill_Records("RSTSSPCX", SEASON_CODE)
        Sort_grdColumns(grdRSTSSPCX, "CUST_CODE,HC_CODE")
        grdRSTSSPCX.Text = "Stock & Sales Plans for " & SEASON_CODE
    End Sub

    Sub XLS_Validation(isTemplate As Boolean)
        Dim sheet_valid As Boolean = True
        Dim sheet_error_msg As String = ""
        XLSR.Clear()

        'If isTemplate Then Delete_Sheets.Clear()
        Delete_Sheets.Clear()

        xls_COLLECTION_CODEs.Clear()

        Dim LINE_TAGsAll As List(Of String) = {"PSIHDG", "TOAHDG", "MOSHDG", "ORGHDG", "LYAHDG", "CUST_CODE", "COLLECTION_CODE", "SEASON_CODE"}.ToList

        For Each LINE_TAG As String In LINE_TAGs
            LINE_TAGsAll.Add(LINE_TAG)
        Next

        For Each LINE_TAG As String In New String() {"TYPGRSBPLN", "TYPGRSPPLN", "TYPADJPLN"}
            LINE_TAGsAll.Add(LINE_TAG)
        Next

        workbook = WorkbookView1.ActiveWorkbook

        If workbook Is Nothing OrElse workbook.Worksheets.Count < 3 OrElse workbook.Worksheets(0).Name <> "Total" Then
            sheet_error_msg = "Workbook does not contain at least 3 sheets beginning with Total Sheet"
        Else
            If isTemplate Then
                Delete_Sheets.Add(workbook.Worksheets(1).Name)
                Delete_Sheets.Add(workbook.Worksheets(2).Name)
            End If

            For i As Integer = 0 To workbook.Worksheets.Count - 1
                Dim sheet_name As String = workbook.Worksheets(i).Name
                If sheet_name <> "Total" And Not Delete_Sheets.Contains(sheet_name) Then
                    xls_COLLECTION_CODEs.Add(sheet_name)
                End If
            Next
            worksheet = workbook.Worksheets("Total")
            For i As Integer = 0 To 999
                Dim LINE_TAG As String = worksheet.Cells(i, 0).Value & ""
                If LINE_TAGsAll.Contains(LINE_TAG) Then
                    XLSR.Add(LINE_TAG, i)
                    LINE_TAGsAll.Remove(LINE_TAG)
                End If
            Next
            If LINE_TAGsAll.Count >= 0 Then
                sheet_error_msg = "Total Worksheet does not have Line Tags for " & Join(LINE_TAGsAll.ToArray, ",")
            End If
        End If

        If sheet_error_msg = "" Then

        End If

        sheet_valid = (sheet_error_msg = "")
    End Sub

    Sub XLS_Refresh_COLLECTION(Optional single_COLLECTION_CODE As String = "")

        Dim rCOLLECTION_CODE As Integer = XLSR("COLLECTION_CODE")

        If single_COLLECTION_CODE <> "" Then
            WorkbookView1.GetLock()
            ASCMAIN1.sql = "Select ICTCOLL1.* from ICTCOLL1 where COLLECTION_CODE = '" & single_COLLECTION_CODE & "'"
        Else
            ASCMAIN1.sql = "Select ICTCOLL1.* from ICTCOLL1,ICTBRAN1" & vbCrLf _
                    & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                    & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_STATUS = 'A' and ICTBRAN1.BRAND_STATUS = 'A'"

            If rowSOTTCLS1.Item("AUTH_REQD") & "" = "1" Then
                ASCMAIN1.sql &= "" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE in (" & vbCrLf

                ' 08/07/18 EMAIL FROM JK WHERE NOTHING SHOWED UP SINCE THE  ENDING DATE OF THE PRIOR SEASON WAS LATER THAN TODAY
                ' TAKING OUT THAT IF STATEMENT SO THAT THE COLLECTIONS SHOW

                'If Format(DTE2, "yyyyMMdd") > Format(Now.Date, "yyyyMMdd") Then ' If SEASON_is_current_or_future Then
                ASCMAIN1.sql &= "" & vbCrLf _
                    & " (Select Distinct ICTCOLL1.COLLECTION_CODE from SATAUTH1,ICTCOLL1" & vbCrLf _
                    & "   where SATAUTH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "     and ICTCOLL1.HC_CODE = SATAUTH1.HC_CODE" & vbCrLf _
                    & "     and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL and SATAUTH1.OPS_YYYYPP_CLOSED IS NULL)" & vbCrLf _
                    & " UNION " & vbCrLf
                'End If

                ASCMAIN1.sql &= "" & vbCrLf _
                    & " (Select Distinct COLLECTION_CODE from RSTSSPC2" & vbCrLf _
                    & " where SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                    & "   and HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "   and CUST_CODE = '" & CUST_CODE & "')" _
                    & " )"

            End If
        End If

        Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(workbook.Worksheets.Count - 1)
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "BRAND_CODE,COLLECTION_CODE")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
            If Not xls_COLLECTION_CODEs.Contains(COLLECTION_CODE) Then
                Dim worksheet3 As SpreadsheetGear.IWorksheet = worksheet2
                For i As Integer = 1 To workbook.Worksheets.Count - 1
                    Dim sheet_name As String = workbook.Worksheets(i).Name
                    If Not Delete_Sheets.Contains(sheet_name) Then
                        If sheet_name > COLLECTION_CODE Then
                            worksheet3 = workbook.Worksheets(i)
                            Exit For
                        End If
                    End If
                Next
                worksheet = workbook.Worksheets.AddBefore(worksheet3)
                worksheet.Name = COLLECTION_CODE
                worksheet.Cells("E4").Activate()
                worksheet.WindowInfo.FreezePanes = True

                worksheet2.UsedRange.Copy(worksheet.Range("A1"), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                For C As Integer = 0 To 50
                    worksheet.Cells(0, C).EntireColumn.ColumnWidth = worksheet2.Cells(0, C).EntireColumn.ColumnWidth
                Next

                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                worksheet.Cells(rCOLLECTION_CODE, 1).Value = COLLECTION_CODE
                worksheet.Cells(rCOLLECTION_CODE, 2).Value = rowICTCOLL1.Item("COLLECTION_NAME")

                xls_COLLECTION_CODEs.Add(COLLECTION_CODE)
            End If
        Next

        If single_COLLECTION_CODE <> "" Then
            WorkbookView1.ReleaseLock()
            Absx1.txtFor("COLLECTION_CODE").Text = ""
        End If
    End Sub

    Sub XLS_Refresh_Allocations()

        ASCMAIN1.sql = sqlSOTALLOX & vbCrLf _
            & "   and SOTALLO2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_START between '" & Format(DTE1.AddMonths(-1), "dd-MMM-yyyy") & "' and '" & Format(DTE2.AddMonths(-1), "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE in ('" & Join(xls_COLLECTION_CODEs.ToArray, "','") & "')"

 
        ASCMAIN1.sql &= " UNION " & vbCrLf _
            & Replace(sqlSOTALLOX, "and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO", "and SOTALLO2.ALLO_CTL_NO (+) = SOTALLO1.ALLO_CTL_NO") & vbCrLf _
            & "   and SOTALLO2.CUST_CODE (+) = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTALLO1.ALLO_CTL_NO in (Select Distinct LINE_KEY from RSTSSPC2" & vbCrLf _
            & " where RSTSSPC2.SEASON_CODE = '" & SEASON_CODE & "' and RSTSSPC2.CUST_CODE = '" & CUST_CODE & "' and LINE_TAG in ('TOAITM','PSIITM'))" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE in ('" & Join(xls_COLLECTION_CODEs.ToArray, "','") & "')"

        Fill_Records("SOTALLOX", "", True, ASCMAIN1.sql)

        worksheet = workbook.Worksheets(0)

        Dim rTOAHDG As Integer = XLSR("TOAHDG")
        Dim rTOAITM As Integer = rTOAHDG + 2
        For Each row As DataRow In dst.Tables("SOTALLOX").Select("ITEM_SNU_CODE = 'N'", "ITEM_CODE,DATE_START")
            range = worksheet.Range(rTOAITM, 0).EntireRow
            rTOAITM += 1
            worksheet.Range(rTOAITM, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
            range.Copy(worksheet.Range(rTOAITM, 0), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
            worksheet.Cells(rTOAITM, 1).Value = row.Item("ALLO_CTL_NO")   ' row.Item("ITEM_CODE")   'B
            worksheet.Cells(rTOAITM, 2).Value = row.Item("COLLECTION_CODE")     'C
            worksheet.Cells(rTOAITM, 3).Value = row.Item("ITEM_DESC")   'D
            worksheet.Cells(rTOAITM, 13).Value = row.Item("QTY_ALLO")   'N
            worksheet.Cells(rTOAITM, 14).Value = row.Item("DATE_START") 'O
            worksheet.Cells(rTOAITM, 15).Value = row.Item("ITEM_CODE")  'P

            For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
                Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
                Dim range2 As SpreadsheetGear.IRange = worksheet2.Range(rTOAITM - 1, 0).EntireRow
                worksheet2.Range(rTOAITM, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                range2.Copy(worksheet2.Range(rTOAITM, 0), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
            Next
        Next

        ' NO NEED TO ADJUST LINE_TAG REFERENCES AFTER TOAITMs because they are already at the bottom

        Dim rPSIHDG As Integer = XLSR("PSIHDG")
        Dim rPSIITM As Integer = rPSIHDG + 1
        For Each row As DataRow In dst.Tables("SOTALLOX").Select("ITEM_SNU_CODE = 'S'", "ITEM_CODE,DATE_START")
            range = worksheet.Range(rPSIITM, 0).EntireRow
            rPSIITM += 1
            worksheet.Range(rPSIITM, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
            range.Copy(worksheet.Range(rPSIITM, 0), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
            worksheet.Cells(rPSIITM, 1).Value = row.Item("ALLO_CTL_NO")   'row.Item("ITEM_CODE")   'B
            worksheet.Cells(rPSIITM, 2).Value = row.Item("COLLECTION_CODE")     'C
            worksheet.Cells(rPSIITM, 3).Value = row.Item("ITEM_DESC")   'D
            worksheet.Cells(rPSIITM, 13).Value = row.Item("QTY_ALLO")   'N
            worksheet.Cells(rPSIITM, 16).Value = row.Item("ITEM_RETAIL_PRICE")      'Q
            worksheet.Cells(rPSIITM, 17).Value = row.Item("DATE_START") 'R
            worksheet.Cells(rPSIITM, 18).Value = row.Item("ITEM_CODE")  'S
            worksheet.Cells(rPSIITM, 19).Value = row.Item("ITEM_CODE_COMPARE_TO")   'T
            worksheet.Cells(rPSIITM, 20).Value = row.Item("ITEM_CODE_COMPARE_TO_ALT")   'U

            For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
                Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
                Dim range2 As SpreadsheetGear.IRange = worksheet2.Range(rPSIITM - 1, 0).EntireRow
                worksheet2.Range(rPSIITM, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                range2.Copy(worksheet2.Range(rPSIITM, 0), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
            Next
        Next


        worksheet = workbook.Worksheets(0)

        Dim XLSR_orig As New Dictionary(Of String, Integer)
        For Each LINE_TAG As String In XLSR.Keys
            XLSR_orig.Add(LINE_TAG, XLSR(LINE_TAG))
        Next

        For Each LINE_TAG As String In XLSR_orig.Keys
            If XLSR(LINE_TAG) > rPSIHDG Then

                Dim OLD_LINE As Integer = XLSR(LINE_TAG)
                XLSR(LINE_TAG) += rPSIITM - (rPSIHDG + 1)

                Dim sLINE As Integer = XLSR(LINE_TAG)
                Dim eLINE As Integer = XLSR(LINE_TAG)

                If LINE_TAG.StartsWith("TOA") Then
                    Get_Range("TOAITM")
                    sLINE = sLINE_TAG
                    eLINE = eLINE_TAG
                    OLD_LINE += 2
                End If

                'If ASCMAIN1.Running_in_VS And LINE_TAG = "TYPADJT" Then MsgBox(LINE_TAG & vbCrLf & workbook.Worksheets("Total").Cells("E63").Formula & vbCrLf & workbook.Worksheets("Total").Cells("E64").Formula)
                'If ASCMAIN1.Running_in_VS And LINE_TAG = "TYPADJT" Then Stop
                For iLine As Integer = sLINE To eLINE
                    OLD_LINE += 1
                    For M As Integer = 0 To 8
                        Dim F As String = worksheet.Cells(iLine, c0 + M).Formula
                        ' IF WE ALLOW THIS LINE TO EXECUTE FOR AAFES, WE GET CIRCULAR REFERENCE
                        If F.StartsWith("=SUM(") And F.EndsWith(CStr(OLD_LINE) & ")") And LINE_TAG <> "TYPADJT" And LINE_TAG <> "TYPGRSPT" Then
                            'THIS SECTION APPEARS TO BE USED ONLY BY THE 2 LINE TAGS THAT I HAVE NOW PREVENTED FROM USING IT
                            'Debug.Print(LINE_TAG)
                            F = Replace(F, CStr(OLD_LINE) & ")", CStr(OLD_LINE + rPSIITM - (rPSIHDG + 1)) & ")")
                            worksheet.Cells(iLine, c0 + M).Formula = F
                        End If
                    Next
                Next

                ' If ASCMAIN1.Running_in_VS And LINE_TAG = "TYPADJT" Then MsgBox(LINE_TAG & vbCrLf & workbook.Worksheets("Total").Cells("E63").Formula & vbCrLf & workbook.Worksheets("Total").Cells("E64").Formula)

            End If
        Next

        XLS_Allocation_Lines.Clear()

        Get_Range("PSIITM")
        For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
            Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
            For i As Integer = sLINE_TAG To eLINE_TAG
                worksheet2.Cells(i, 0).EntireRow.Hidden = (CStr(worksheet2.Cells(i, 2).Value & "") <> CStr(worksheet2.Cells(1, 1).Value & ""))
                If COLLECTION_CODE = xls_COLLECTION_CODEs(0) And i > sLINE_TAG Then XLS_Allocation_Lines.Add(worksheet2.Cells(i, 1).Value, i)
            Next
        Next

        ' sLINE_TAG and eLINE_TAG are actually Sample Items
        For Each i As Integer In New Integer() {sLINE_TAG, eLINE_TAG}
            worksheet.Cells(i, 0).EntireRow.Hidden = True
            '    If worksheet.Cells(i, 3).Value = "Sample Item" Then Stop
            For m As Integer = 0 To 6
                worksheet.Cells(i, c0 + m).Formula = ""
            Next
        Next


        Get_Range("TOAITM")
        For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
            Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
            For i As Integer = sLINE_TAG To eLINE_TAG
                worksheet2.Cells(i, 0).EntireRow.Hidden = (CStr(worksheet2.Cells(i, 2).Value & "") <> CStr(worksheet2.Cells(1, 1).Value & ""))
                If COLLECTION_CODE = xls_COLLECTION_CODEs(0) And i > sLINE_TAG Then XLS_Allocation_Lines.Add(worksheet2.Cells(i, 1).Value, i)
            Next
        Next
    End Sub

    Sub Get_Range(LINE_TAG As String)
        sLINE_TAG = -1
        eLINE_TAG = -1

        Dim i As Integer = 0
        Do Until sLINE_TAG <> -1 And eLINE_TAG < i - 1
            If sLINE_TAG = -1 Then
                If workbook.Worksheets(0).Cells(i, 0).Value & "" = LINE_TAG Then
                    sLINE_TAG = i
                End If
            End If
            If workbook.Worksheets(0).Cells(i, 0).Value & "" = LINE_TAG Then
                eLINE_TAG = i
            End If
            i += 1
        Loop
    End Sub

    Sub Get_TY_Actuals()

        Dim XLR As Int64 = 0
        Dim sql As String = ""

        ' TYA SLS (RETAIL)
        sql = ""
        For I As Integer = 0 To 6
            sql &= ", SUM(DECODE(RSTRETL1.OPS_YYYYPP,'" & YPs(I) & "',RSTRETL1.AMT_SOLD,0)) TYPSLSACT_" & CStr(I) & vbCrLf
        Next

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & sql _
            & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYPP between '" & YPs(0) & "' and '" & YPs(6) & "'" & vbCrLf _
            & "   group by ICTCOLL1.COLLECTION_CODE"
        Dim tblTYASLS As DataTable = ASCDATA1.GetDataTable
        tblTYASLS.PrimaryKey = New DataColumn() {tblTYASLS.Columns("COLLECTION_CODE")}


        ' TYA GRS & ADJ (@RETAIL)

        Dim Z As String = "THEN SOTINVH2.ORDR_QTY_SHIP * TRUNC(10000 * DECODE(NVL(SOTINVH2.ITEM_RETAIL_PRICE,0),0,NVL(SOTINVH2.ORDR_UNIT_PRICE,0)*100/60,NVL(SOTINVH2.ITEM_RETAIL_PRICE,0)))/10000 ELSE 0 END"
        Dim W As String = "THEN SOTINVH2.ORDR_QTY_SHIP * TRUNC(10000 * NVL(SOTINVH2.ORDR_UNIT_PRICE,0))/10000 ELSE 0 END"
        sql = ""
        For I As Integer = 0 To 6
            Dim Y As String = "SOTINVH2.ORDR_YYYYPP_UPDATED='" & YPs(I) & "'"
            sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' AND ICTITEM1.ITEM_BASIC_PROMO = 'B' " & Z & ") TYAGRSB_" & CStr(I) & vbCrLf
            sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' AND ICTITEM1.ITEM_BASIC_PROMO = 'P' " & Z & ") TYAGRSP_" & CStr(I) & vbCrLf
            sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'C' " & Z & ") TYAADJ_" & CStr(I) & vbCrLf
            sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' AND ICTITEM1.ITEM_BASIC_PROMO = 'B' " & W & ") TYPGRSBNET_" & CStr(I) & vbCrLf
            sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' AND ICTITEM1.ITEM_BASIC_PROMO = 'P' " & W & ") TYPGRSPNET_" & CStr(I) & vbCrLf
        Next

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & sql _
            & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(0) & "' and '" & YPs(6) & "'" & vbCrLf _
            & "   group by ICTCOLL1.COLLECTION_CODE"
        Dim tblTYA As DataTable = ASCDATA1.GetDataTable
        tblTYA.PrimaryKey = New DataColumn() {tblTYA.Columns("COLLECTION_CODE")}

        ' TYA Open + Pick

        Z = "THEN (NVL(SOTORDR2.ORDR_QTY_OPEN,0)+NVL(SOTORDR2.ORDR_QTY_PICK,0)) * DECODE(NVL(SOTORDR2.ITEM_RETAIL_PRICE,0),0,TRUNC(10000 * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)*100/60)/10000,NVL(SOTORDR2.ITEM_RETAIL_PRICE,0)) ELSE 0 END"
        W = "THEN (NVL(SOTORDR2.ORDR_QTY_OPEN,0)+NVL(SOTORDR2.ORDR_QTY_PICK,0)) * TRUNC(10000 * NVL(SOTORDR2.ORDR_UNIT_PRICE,0))/10000 ELSE 0 END"
        sql = ""
        Dim LTE As String = "<="

        LTE = "="
        If YPs(0) <= ASCMAIN1.CYP And YPs(6) >= ASCMAIN1.CYP Then
            LTE = "<="
        End If

        For I As Integer = 0 To 6
            If YPs(I) < ASCMAIN1.CYP Then
                sql &= ", 0 TYAGRSB_" & CStr(I) & vbCrLf
                sql &= ", 0 TYAGRSP_" & CStr(I) & vbCrLf
                sql &= ", 0 TYAADJ_" & CStr(I) & vbCrLf
                sql &= ", 0 TYPGRSBNET_" & CStr(I) & vbCrLf
                sql &= ", 0 TYPGRSPNET_" & CStr(I) & vbCrLf
            Else
                Dim Y As String = "TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') " & LTE & " '" & YPs(I) & "'"
                sql &= ", SUM(CASE WHEN " & Y & " AND ICTITEM1.ITEM_BASIC_PROMO = 'B' " & Z & ") TYAGRSB_" & CStr(I) & vbCrLf
                sql &= ", SUM(CASE WHEN " & Y & " AND ICTITEM1.ITEM_BASIC_PROMO = 'P' " & Z & ") TYAGRSP_" & CStr(I) & vbCrLf
                sql &= ", 0 TYAADJ_" & CStr(I) & vbCrLf
                sql &= ", SUM(CASE WHEN " & Y & " AND ICTITEM1.ITEM_BASIC_PROMO = 'B' " & Z & ") TYPGRSBNET_" & CStr(I) & vbCrLf
                sql &= ", SUM(CASE WHEN " & Y & " AND ICTITEM1.ITEM_BASIC_PROMO = 'P' " & Z & ") TYPGRSPNET_" & CStr(I) & vbCrLf
                LTE = "="
            End If
        Next

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & sql _
            & " from SOTORDR2,SOTORDR1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
            & "   and SOTORDR1.ORDR_STATUS >= 'O' and SOTORDR1.ORDR_STATUS <= 'P'" & vbCrLf _
            & "   group by ICTCOLL1.COLLECTION_CODE"

        Dim tblOP As DataTable = ASCDATA1.GetDataTable
        tblOP.PrimaryKey = New DataColumn() {tblOP.Columns("COLLECTION_CODE")}


        WorkbookView1.GetLock()

        For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
            If ws.ProtectContents Then
                ws.Unprotect(XLS_PWD)
            End If

            ' Clear out Actuals in case Open Order amount has moved, or history was corrected
            For I As Integer = 0 To 6
                ws.Cells(XLSR("TYAGRSB"), c0 + I).Value = 0
                ws.Cells(XLSR("TYAGRSP"), c0 + I).Value = 0
                ws.Cells(XLSR("TYAADJ"), c0 + I).Value = 0
            Next

            Dim rowTYASLS As DataRow = tblTYASLS.Rows.Find(COLLECTION_CODE)
            If rowTYASLS IsNot Nothing Then
                For I As Integer = 0 To 6
                    ws.Cells(XLSR("TYPSLSACT"), c0 + I).Value = Val(rowTYASLS.Item("TYPSLSACT_" & CStr(I)) & "") / 1000
                Next
            End If

            Dim rowTYA As DataRow = tblTYA.Rows.Find(COLLECTION_CODE)
            Dim rowOP As DataRow = tblOP.Rows.Find(COLLECTION_CODE)
            If rowTYA IsNot Nothing Or rowOP IsNot Nothing Then
                For I As Integer = 0 To 6
                    Dim A(5) As Decimal
                    If rowTYA IsNot Nothing Then
                        A(1) += Val(rowTYA.Item("TYAGRSB_" & CStr(I)) & "") / 1000
                        A(2) += Val(rowTYA.Item("TYAGRSP_" & CStr(I)) & "") / 1000
                        A(3) += Val(rowTYA.Item("TYAADJ_" & CStr(I)) & "") / 1000
                        A(4) += Val(rowTYA.Item("TYPGRSBNET_" & CStr(I)) & "") / 1000
                        A(5) += Val(rowTYA.Item("TYPGRSPNET_" & CStr(I)) & "") / 1000
                    End If
                    If rowOP IsNot Nothing Then
                        A(1) += Val(rowOP.Item("TYAGRSB_" & CStr(I)) & "") / 1000
                        A(2) += Val(rowOP.Item("TYAGRSP_" & CStr(I)) & "") / 1000
                        A(3) += Val(rowOP.Item("TYAADJ_" & CStr(I)) & "") / 1000
                        A(4) += Val(rowOP.Item("TYPGRSBNET_" & CStr(I)) & "") / 1000
                        A(5) += Val(rowOP.Item("TYPGRSPNET_" & CStr(I)) & "") / 1000
                    End If
                    'ws.Cells(XLSR("TYAGRSB"), c0 + I).Value = Val(rowTYA.Item("TYAGRSB_" & CStr(I)) & "") / 1000
                    'ws.Cells(XLSR("TYAGRSP"), c0 + I).Value = Val(rowTYA.Item("TYAGRSP_" & CStr(I)) & "") / 1000
                    'ws.Cells(XLSR("TYAADJ"), c0 + I).Value = Val(rowTYA.Item("TYAADJ_" & CStr(I)) & "") / 1000
                    ws.Cells(XLSR("TYAGRSB"), c0 + I).Value = A(1)
                    ws.Cells(XLSR("TYAGRSP"), c0 + I).Value = A(2)
                    ws.Cells(XLSR("TYAADJ"), c0 + I).Value = A(3)
                    If YPs(I) < ASCMAIN1.CYP Then ' IF MONTH IS AN ACTUAL MONTH
                        ws.Cells(XLSR("TYPGRSBNET"), c0 + I).Value = A(4)
                        ws.Cells(XLSR("TYPGRSPNET"), c0 + I).Value = A(5)
                        ' ws.Cells(XLSR("TYPGRSNET"), c0 + I).Formula = "=" & ws.Cells(XLSR("TYPGRSBNET"), c0 + I).FormulaR1C1 & "+" & ws.Cells(XLSR("TYPGRSPNET"), c0 + I).FormulaR1C1
                        ws.Cells(XLSR("TYPGRSNET"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPGRSBNET"), c0 + I) & "+" & Excel_Cell0(XLSR("TYPGRSPNET"), c0 + I)
                    End If
                Next
            End If

            For I As Integer = 0 To 6
                If YPs(I) < ASCMAIN1.CYP Then
                    ws.Cells(XLSR("TYPSLSACTZ"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPSLSACT"), c0 + I)
                    ws.Cells(XLSR("TYPGRSB"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYAGRSB"), c0 + I)
                    ws.Cells(XLSR("TYPGRSP"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYAGRSP"), c0 + I)
                    ws.Cells(XLSR("TYPADJ"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYAADJ"), c0 + I)
                Else
                    ws.Cells(XLSR("TYPSLSACTZ"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPSLS"), c0 + I)
                    ws.Cells(XLSR("TYPGRSB"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPGRSBT"), c0 + I)
                    ws.Cells(XLSR("TYPGRSP"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPGRSPT"), c0 + I)
                    ws.Cells(XLSR("TYPADJ"), c0 + I).Formula = "=" & Excel_Cell0(XLSR("TYPADJT"), c0 + I)
                End If
            Next
        Next

        WorkbookView1.ReleaseLock()

    End Sub

    Sub Get_LY_Actuals()

        ' This method should be called only once, when EntryMode = "N"

        Dim sql As String = ""

        ' LYA BOM

        sql = ""
        For I As Integer = 0 To 6
            sql &= ", SUM (RSTSSPC2.AMT_" & CStr(I) & ") LYABOM_" & CStr(I) & vbCrLf
        Next

        ASCMAIN1.sql = "Select RSTSSPC2.COLLECTION_CODE" & vbCrLf _
            & sql _
            & " from RSTSSPC2" & vbCrLf

        If EntryMode = "N" Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS
            ASCMAIN1.sql &= "" _
                & " where RSTSSPC2.LINE_TAG = 'TYPBOM'" & vbCrLf _
                & "   and RSTSSPC2.SEASON_CODE = '" & SEASON_CODE_LY & "'" & vbCrLf _
                & "   and RSTSSPC2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.HC_CODE = '" & HC_CODE & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= "" _
                & " where RSTSSPC2.LINE_TAG = 'LYABOM'" & vbCrLf _
                & "   and RSTSSPC2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.HC_CODE = '" & HC_CODE & "'" & vbCrLf
        End If

        ASCMAIN1.sql &= "   group by RSTSSPC2.COLLECTION_CODE"

        Dim tblLYABOM As DataTable = ASCDATA1.GetDataTable

        If EntryMode = "N" And tblLYABOM.Rows.Count = 0 Then
            ASCMAIN1.sql = "Select ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", SUM (DECODE(GLTPARM3.REL_WEEK,GLTPARM3.MAX_WEEK,QTY_EOW,0) * ICTITEM1.ITEM_RETAIL_PRICE)/1000 LYABOM_0" & vbCrLf _
                & ", 0 LYABOM_1, 0 LYABOM_2, 0 LYABOM_3, 0 LYABOM_4, 0 LYABOM_5, 0 LYABOM_6" & vbCrLf _
                & "from ICTITEM1,RSTRETL1,GLTPARM3,ICTCOLL1" & vbCrLf _
                & "where ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "  and RSTRETL1.OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(YPs(0), -13) & "'" & vbCrLf _
                & "  and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf _
                & "  and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "  and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "  and RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "group by ICTITEM1.COLLECTION_CODE"
            tblLYABOM = ASCDATA1.GetDataTable
        End If

        tblLYABOM.PrimaryKey = New DataColumn() {tblLYABOM.Columns("COLLECTION_CODE")}

        ' LYA SLS (RETAIL)

        If EntryMode = "N" Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS

            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM(DECODE(RSTRETL1.OPS_YYYYPP,'" & ASCMAIN1.Period_Calc(YPs(I), -12) & "',RSTRETL1.AMT_SOLD,0)) LYASLS_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE" & vbCrLf _
                & sql _
                & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   and RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTRETL1.OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(YPs(0), -12) & "' and '" & ASCMAIN1.Period_Calc(YPs(6), -12) & "'" & vbCrLf _
                & "   group by ICTCOLL1.COLLECTION_CODE"
        Else
            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM (1000 * RSTSSPC2.AMT_" & CStr(I) & ") LYASLS_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select RSTSSPC2.COLLECTION_CODE" & vbCrLf _
                & sql _
                & " from RSTSSPC2" & vbCrLf _
                & " where RSTSSPC2.LINE_TAG = 'LYASLS'" & vbCrLf _
                & "   and RSTSSPC2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   group by RSTSSPC2.COLLECTION_CODE"
        End If

        Dim tblLYASLS As DataTable = ASCDATA1.GetDataTable
        tblLYASLS.PrimaryKey = New DataColumn() {tblLYASLS.Columns("COLLECTION_CODE")}

        ' LYA GRS & ADJ (@RETAIL)

        If EntryMode = "N" Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS
            Dim Z As String = "THEN SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_RETAIL_PRICE ELSE 0 END"
            sql = ""
            For I As Integer = 0 To 6
                Dim Y As String = "SOTINVH2.ORDR_YYYYPP_UPDATED='" & ASCMAIN1.Period_Calc(YPs(I), -12) & "'"
                sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' " & Z & ") LYAGRS_" & CStr(I) & vbCrLf
                sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'C' " & Z & ") LYAADJ_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE" & vbCrLf _
                & sql _
                & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & ASCMAIN1.Period_Calc(YPs(0), -12) & "' and '" & ASCMAIN1.Period_Calc(YPs(6), -12) & "'" & vbCrLf _
                & "   group by ICTCOLL1.COLLECTION_CODE"
        Else
            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM (DECODE(RSTSSPC2.LINE_TAG,'LYAGRS',1000 * RSTSSPC2.AMT_" & CStr(I) & ",0)) LYAGRS_" & CStr(I) & vbCrLf
                sql &= ", SUM (DECODE(RSTSSPC2.LINE_TAG,'LYAADJ',1000 * RSTSSPC2.AMT_" & CStr(I) & ",0)) LYAADJ_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select RSTSSPC2.COLLECTION_CODE" & vbCrLf _
                & sql _
                & " from RSTSSPC2" & vbCrLf _
                & " where RSTSSPC2.LINE_TAG in ('LYAGRS','LYAADJ')" & vbCrLf _
                & "   and RSTSSPC2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTSSPC2.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   group by RSTSSPC2.COLLECTION_CODE"
        End If

        Dim tblLYA As DataTable = ASCDATA1.GetDataTable
        tblLYA.PrimaryKey = New DataColumn() {tblLYA.Columns("COLLECTION_CODE")}


        WorkbookView1.GetLock()

        For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)

            Dim rowLYABOM As DataRow = tblLYABOM.Rows.Find(COLLECTION_CODE)
            If rowLYABOM IsNot Nothing Then
                For I As Integer = 0 To 0 '6
                    ws.Cells(XLSR("LYABOM"), 4 + I).Value = Val(rowLYABOM.Item("LYABOM_" & CStr(I)) & "")
                Next
            End If

            Dim rowLYASLS As DataRow = tblLYASLS.Rows.Find(COLLECTION_CODE)
            If rowLYASLS IsNot Nothing Then
                For I As Integer = 0 To 6
                    ws.Cells(XLSR("LYASLS"), 4 + I).Value = Val(rowLYASLS.Item("LYASLS_" & CStr(I)) & "") / 1000
                Next
            End If

            Dim rowLYA As DataRow = tblLYA.Rows.Find(COLLECTION_CODE)
            If rowLYA IsNot Nothing Then
                For I As Integer = 0 To 6
                    ws.Cells(XLSR("LYAGRS"), 4 + I).Value = Val(rowLYA.Item("LYAGRS_" & CStr(I)) & "") / 1000
                    ws.Cells(XLSR("LYAADJ"), 4 + I).Value = Val(rowLYA.Item("LYAADJ_" & CStr(I)) & "") / 1000
                Next
            End If
        Next

        WorkbookView1.ReleaseLock()

    End Sub

    Private Sub cmdLockOriginal_Click(sender As Object, e As EventArgs) Handles cmdLockOriginal.Click

        WorkbookView1.GetLock()

        For Each COLLECTION_CODE As String In xls_COLLECTION_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(COLLECTION_CODE)
            ws.Unprotect(XLS_PWD)
            ws.Cells(XLSR("TYOBOM"), c0 + 0).Value = ws.Cells(XLSR("TYPBOM"), c0 + 0).Value

            For i As Integer = 0 To 6
                ws.Cells(XLSR("TYOSLS"), c0 + i).Value = ws.Cells(XLSR("TYPSLS"), c0 + i).Value
                ws.Cells(XLSR("TYOGRS"), c0 + i).Value = ws.Cells(XLSR("TYPGRS"), c0 + i).Value
                ws.Cells(XLSR("TYOADJ"), c0 + i).Value = ws.Cells(XLSR("TYPADJ"), c0 + i).Value
            Next
            ws.Protect(XLS_PWD)
        Next

        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub chkShowLINE_TAGs_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowLINE_TAGs.CheckedChanged

        If Not ScreenMode Then Exit Sub

        WorkbookView1.GetLock()

        For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
            If worksheet.ProtectContents Then
                worksheet.Unprotect(XLS_PWD)
            End If
            worksheet.Range("A1:C1").EntireColumn.Hidden = Not chkShowLINE_TAGs.Checked
            worksheet.Protect(XLS_PWD)
        Next

        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub grdRSTSSPCX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSSPCX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value & ""
            Absx1.txtFor("HC_CODE").Text = e.Row.Cells("HC_CODE").Value & ""
            Click_Command("Edit")
        End If
    End Sub

    Sub Box_Range()
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
    End Sub
    Sub Set_Data_Block(ByRef r As Integer, COLLECTION_CODE As String, tbl As DataTable, CHANNEL_CODE As String)

        r += 2

        If COLLECTION_CODE = "" Then
            worksheet.Cells(r, c0 - 1).Value = "Total"
        Else
            worksheet.Cells(r, c0 - 1).Value = COLLECTION_CODE
        End If
        For i As Integer = 0 To 6 + 2
            worksheet.Cells(r, c0 + i).Formula = "=Total!" & Excel_Cell0(2, c0 + i)
        Next
        range = worksheet.Cells(r, c0 - 1, r, c0 + 6 + 2)
        range.Font.Bold = True
        Box_Range()

        If SEASON_TYPE = "F" Then

            worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = SEASON_CODE_prior & " 445"
            worksheet.Cells(r, c0 + 6 + 2 + 1 + 2).Value = "Total Year"

            range = worksheet.Cells(r, c0 + 6 + 2 + 1 + 1, r, c0 + 6 + 2 + 1 + 2)
            range.Font.Bold = True
            Box_Range()
        End If

        Dim r0 As Integer = r

        r += 1 : Set_Data_Block_row(r, 7, COLLECTION_CODE)

        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Financial Retail Plan" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Blue
        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Carryover" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Red

        r += 1 : Set_Data_Block_row(r, 9, COLLECTION_CODE)
        r += 1 : Set_Data_Block_row(r, 10, COLLECTION_CODE)
        r += 1 : Set_Data_Block_row(r, 11, COLLECTION_CODE)
        r += 1 : Set_Data_Block_row(r, 16, COLLECTION_CODE) : worksheet.Cells(r, c0 - 1, r, c0 + 6 + 2).Interior.Color = SpreadsheetGear.Colors.LightGray
        If SEASON_TYPE = "F" Then
            worksheet.Cells(r, c0 + 6 + 2 + 1 + 1, r, c0 + 6 + 2 + 1 + 2).Interior.Color = SpreadsheetGear.Colors.LightGray
        End If

        If CHANNEL_CODE = "" Then
        Else
            r += 1 : worksheet.Cells(r, c0 - 1).Value = "Financial Gross Plan" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Blue

            Dim BUD_PS As Decimal = 0
            Dim FBUD(6) As Decimal
            If COLLECTION_CODE = "Total" Then
                For i As Integer = 0 To 6
                    FBUD(i) = Val(tbl.Compute("SUM(P" & Format(i, "0") & ")", "CHANNEL_CODE='1'") & "") / 1000
                Next
            Else
                Dim row As DataRow = tbl.Rows.Find(New String() {"1", COLLECTION_CODE})
                If row IsNot Nothing Then
                    For i As Integer = 0 To 6
                        FBUD(i) = Val(row.Item("P" & Format(i, "0")) & "") / 1000
                    Next
                    If SEASON_TYPE = "F" Then
                        BUD_PS = Val(row.Item("BUD_PS") & "")
                    End If
                End If
            End If
            For i As Integer = 0 To 6
                worksheet.Cells(r, c0 + i).Value = FBUD(i)
            Next

            worksheet.Cells(r, c0 + 6 + 1).Formula = "=SUM(" & Excel_Cell0(r, c0 + 0) & ":" & Excel_Cell0(r, c0 + 5) & ")"
            worksheet.Cells(r, c0 + 6 + 2).Formula = "=SUM(" & Excel_Cell0(r, c0 + 1) & ":" & Excel_Cell0(r, c0 + 6) & ")"

            If SEASON_TYPE = "F" And COLLECTION_CODE <> "" Then
                worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = BUD_PS / 1000
            End If
        End If

        r += 1 : Set_Data_Block_row(r, 17, COLLECTION_CODE)
        worksheet.Cells(r, c0 + 6 + 1).Value = DBNull.Value  ' EOM
        worksheet.Cells(r, c0 + 6 + 2).Value = DBNull.Value  ' EOM

        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Retail Trend"
        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Shipping %Chg"
        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Original Retail Plan"
        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Original Shipping @Retail"

        If SEASON_TYPE = "F" Then
            Dim m0 As Integer = c0 + 6 + 2 + 1
            worksheet.Cells(r0 + 1, m0 + 2).Formula = "=" & Excel_Cell0(r0 + 1, c0 + 6 + 1) & "+" & Excel_Cell0(r0 + 1, m0 + 1)

            range = worksheet.Cells(Excel_Cell0(r0 + 1, m0 + 2) & ":" & Excel_Cell0(r, m0 + 2))
            worksheet.Cells(r0 + 1, m0 + 2).Copy(range, SpreadsheetGear.PasteType.FormulasAndNumberFormats, SpreadsheetGear.PasteOperation.None, False, False)
            worksheet.Cells(r, m0 + 1).Value = DBNull.Value  ' EOM
            worksheet.Cells(r, m0 + 2).Value = DBNull.Value  ' EOM
        End If

    End Sub

    Sub Set_Data_Block_row(r As Integer, rSource As Integer, COLLECTION_CODE As String)

        Dim LINE_TAG As String = workbook.Worksheets("Total").Cells(rSource, 0).Value
        Dim SHEET_NAME As String = HC_CODE
        If SHEET_NAME = "" Then SHEET_NAME = "Total"

        worksheet.Cells(r, c0 - 1).Formula = "=" & SHEET_NAME & "!" & Excel_Cell0(rSource, c0 - 1)
        For i As Integer = 0 To 6 + 2
            worksheet.Cells(r, c0 + i).Formula = "='" & SHEET_NAME & "'!" & Excel_Cell0(rSource, c0 + i)
        Next

        If SEASON_TYPE = "F" And COLLECTION_CODE <> "" Then
            Dim row As DataRow = dst.Tables("RSTSSPCR").Rows.Find(New String() {COLLECTION_CODE, LINE_TAG})
            If row IsNot Nothing Then
                worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = Val(row.Item("AMT") & "")
            End If
        End If
    End Sub
End Class