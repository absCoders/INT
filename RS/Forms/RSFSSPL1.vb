Imports Infragistics.Win.UltraWinGrid

Public Class RSFSSPL1

    ' talk to SP about limiting this screen 
    '  to certain trade classes - like just those requiring AUths
    ' PROB NEED TO RECONCILE CHANNEL1 TO THOSE WE PREPARE SSPS FOR

    Dim rowRSTSSPL1 As DataRow

    Dim rowARTCUST1 As DataRow
    Dim rowICTSEAS1 As DataRow
    Dim rowSOTTCLS1 As DataRow
    Dim rowSOTPCLS1 As DataRow

    Dim CUST_CODE As String
    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String

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
    Dim xls_HC_CODEs As New List(Of String)
    Dim Delete_Sheets As New List(Of String)
    Dim DTE1 As Date
    Dim DTE2 As Date
    Dim XLS_NO As String
    Dim XLS_PWD As String = "ABS"
    Dim XLS_Allocation_Lines As New Dictionary(Of String, Integer)
    Dim YPs_Imported As New List(Of String)

    Dim RSTSSPL2 As String

    ' LIST OF THE LINE TAGS IN THE ROLL-UP

    Dim LINE_TAGs_Roll_Up As String = "'TYPSLSACTZ','TYPGRSB','TYPGRSP','TYPGRS','TYPGRSBNET','TYPGRSPNET','TYPGRSNET','TYPEOM'"
    'Dim LINE_TAGs_Roll_Up As String = "'TYPSLSACTZ','TYPGRSB','TYPGRSP','TYPGRS','TYPGRSNET','TYPEOM'"

    ' LIST OF LINE TAGS DEFINING THE LINES THAT CONTAIN 7 MOS OF NUMERIC DATA THAT SHOULD BE SAVED WITH SSP
    Dim LINE_TAGs As List(Of String) =
        {"TYPBOM", "TYPSLS", "TYPSLSACT", "TYPSLSACTZ",
         "TYPGRSB", "TYPGRSP", "TYPGRS", "TYPADJ", "TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET", "TYPEOM",
         "TYPSHP%LY", "TYPSLS%LY",
         "CUSEOM", "CUSRTL", "CUSSTC",
         "TYOBOM", "TYOSLS", "TYOGRS", "TYOADJ", "TYOEOM",
         "LYABOM", "LYASLS", "LYAGRS", "LYAADJ", "LYAEOM",
         "TYPB1", "TYPB2", "TYPB3", "TYPB4", "TYPB5", "TYPB6", "TYPB7", "TYPB8", "TYPB9",
         "TYAGRSB", "TYAGRSP", "TYAADJ",
         "TYPGRSBT", "TYPGRSPT", "TYPADJT",
         "TYPDAM", "TYPDIS", "TYPOVR", "TYPSET", "TYPCLS", "PIPE", "CARRYOVR", "RTLOHB"}.ToList

    Dim YPs() As String
    Dim c0 As Integer = 4 ' OFFSET FOR 1ST MONTH OF DATA, COL E = 4

    Dim blnConsolidated As Boolean = False
    Dim CUST_CODES_consolidated As New List(Of String)

    Dim CUST_CODEs_import As String = ""
    Dim ITEM_CODEs_import As String = ""

    Dim filename_special As String = ""
    Dim automated_XLS As Boolean = False
    Dim automated_Customer_HC_Pivot = False

    Dim MONTH_HDGs As New List(Of String)
    Dim updateAsRequired As Boolean = False

    Dim RollUpCustLabel As String = ""
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1, False)
            Create_TDA(.Tables.Add, "RSTSSPL1", "*", 2)
            Create_TDA(.Tables.Add, "RSTSSPL2", "*", 2)
            Create_TDA(.Tables.Add, "RSTSSPL3", "*", 3)
            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)

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

            ASCMAIN1.sql = "Select RSTSSPL1.*,SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                & " from RSTSSPL1,ARTCUST1,SOTTCLS1" & vbCrLf _
                & " where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = RSTSSPL1.CUST_CODE" & vbCrLf _
                & "   and SEASON_CODE = :PARM1"

            Dim YP_LM As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

            ASCMAIN1.sql = $"Select Y.*,SOTTCLS1.CHANNEL_CODE
                , CASE WHEN YP_INIT < LYP_END AND (YP_LAST IS NULL OR YP_LAST < '{YP_LM}') THEN '1' ELSE '0' END UPD
                from (
                Select RSTSSPL1.*
                , PERIOD_CALC(SUBSTR(SEASON_CODE,1,4) || DECODE(SUBSTR(SEASON_CODE,5,1),'F','07','01'),+0) TYP_BEG
                , PERIOD_CALC(SUBSTR(SEASON_CODE,1,4) || DECODE(SUBSTR(SEASON_CODE,5,1),'F','07','01'),+6) TYP_END
                , PERIOD_CALC(SUBSTR(SEASON_CODE,1,4) || DECODE(SUBSTR(SEASON_CODE,5,1),'F','07','01'),+6-12-6) LYP_BEG
                , PERIOD_CALC(SUBSTR(SEASON_CODE,1,4) || DECODE(SUBSTR(SEASON_CODE,5,1),'F','07','01'),+6-12) LYP_END
                 from RSTSSPL1 where SEASON_CODE = :PARM1
                ) Y,ARTCUST1,SOTTCLS1
                where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE 
                  and ARTCUST1.CUST_CODE = Y.CUST_CODE"

            Create_TDA(.Tables.Add, "RSTSSPLX", "**", 0, False, "V", 2)
            .Tables("RSTSSPLX").Columns.Add("SEL")
            .Tables("RSTSSPLX").Columns("SEL").DefaultValue = "0"

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
                .Columns.Add("HC_CODE")
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
                .PrimaryKey = New DataColumn() { .Columns("CHANNEL_CODE"), .Columns("HC_CODE")}
            End With


            With .Tables.Add("RSTSSPP1")
                .Columns.Add("SEASON_CODE")
                .Columns.Add("CUST_CODE")
                .Columns.Add("HC_CODE")
                .Columns.Add("LINE_NO", GetType(System.Int32))
                .Columns.Add("LINE_DESC")
                .Columns.Add("M00", GetType(System.Decimal))
                .Columns.Add("M01", GetType(System.Decimal))
                .Columns.Add("M02", GetType(System.Decimal))
                .Columns.Add("M03", GetType(System.Decimal))
                .Columns.Add("M04", GetType(System.Decimal))
                .Columns.Add("M05", GetType(System.Decimal))
                .Columns.Add("M06", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("SEASON_CODE"), .Columns("CUST_CODE"), .Columns("HC_CODE"), .Columns("LINE_NO")}
            End With


            With .Tables.Add("RSTSSPP2")
                .Columns.Add("SEASON_CODE")
                .Columns.Add("CUST_CODE")
                .Columns.Add("ALLO_CTL_NO")
                .Columns.Add("HC_CODE")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("ITEM_DESC")
                .Columns.Add("ITEM_RETAIL_PRICE", GetType(System.Decimal))
                .Columns.Add("M00", GetType(System.Decimal))
                .Columns.Add("M01", GetType(System.Decimal))
                .Columns.Add("M02", GetType(System.Decimal))
                .Columns.Add("M03", GetType(System.Decimal))
                .Columns.Add("M04", GetType(System.Decimal))
                .Columns.Add("M05", GetType(System.Decimal))
                .Columns.Add("M06", GetType(System.Decimal))
                .Columns.Add("U00", GetType(System.Int32))
                .Columns.Add("U01", GetType(System.Int32))
                .Columns.Add("U02", GetType(System.Int32))
                .Columns.Add("U03", GetType(System.Int32))
                .Columns.Add("U04", GetType(System.Int32))
                .Columns.Add("U05", GetType(System.Int32))
                .Columns.Add("U06", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("SEASON_CODE"), .Columns("CUST_CODE"), .Columns("ALLO_CTL_NO"), .Columns("HC_CODE"), .Columns("ITEM_CODE")}
            End With



            ' NOTE - THE LINE TAGS IN THE WHERE CLAUSE MAY NEED TO BE EXPANDED IF WE EVER ADD LINES TO THE ROLLUP
            '      - SEE LINE_TAGS_Roll_up in Form Declarations
            ASCMAIN1.sql = "Select HC_CODE, LINE_TAG" & vbCrLf _
                & ", SUM (NVL(AMT_0,0)+NVL(AMT_1,0)+NVL(AMT_2,0)+NVL(AMT_3,0)+NVL(AMT_4,0)+NVL(AMT_5,0)) AMT" & vbCrLf _
                & " from RSTSSPL2 " & vbCrLf _
                & " where SEASON_CODE = '" & SEASON_CODE_prior & "'" & vbCrLf _
                & $"   and LINE_TAG IN ({LINE_TAGs_Roll_Up})" & vbCrLf _
                & " group by HC_CODE, LINE_TAG"
            Create_TDA(.Tables.Add, "RSTSSPLR", "**", 0, False, , 2)

            Create_TDA(.Tables.Add, "DPTITMF1", "*")
        End With

        Fill_Records("ICTCOLL0")

        Dim vp As Integer = 0
        grdSATBUDWX.DataSource = dst.Tables("SATBUDWX")
        With grdSATBUDWX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "HC_CODE" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                    gcol.Width = 80
                    gcol.Header.Caption = "HighColl"
                    gcol.Header.Fixed = True
                    Create_Summary(grdSATBUDWX, gcol.Key, "Count")
                ElseIf gcol.Key = "CHANNEL_CODE" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                    gcol.Width = 60
                    gcol.Header.Caption = "Channel"
                    gcol.Header.Fixed = True
                Else
                    gcol.Width = 90
                    gcol.Format = "###,##0"
                    Create_Summary(grdSATBUDWX, gcol.Key)
                    If gcol.Key.StartsWith("R") Then
                        gcol.Header.Appearance.BackColor2 = Color.Gold
                    ElseIf gcol.Key.StartsWith("P") Then
                        gcol.Header.Appearance.BackColor2 = Color.HotPink
                    ElseIf gcol.Key.StartsWith("S") Then
                        gcol.Header.Appearance.BackColor2 = Color.LightGreen
                    ElseIf gcol.Key.StartsWith("F") Then
                        gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    ElseIf gcol.Key.StartsWith("X") Then
                        gcol.Header.Appearance.BackColor2 = Color.Orange
                    End If

                    If gcol.Key.EndsWith("0") Then
                        vp += 1
                        gcol.Header.VisiblePosition = vp
                        gcol.CellAppearance.BackColor = Color.LightGray
                        gcol.Header.Fixed = True
                    End If
                End If
            Next
        End With

        grdRSTSSPL3.DataSource = dst.Tables("RSTSSPL3")

        With grdRSTSSPL3.DisplayLayout.Bands(0)
            For i As Integer = 1 To 6
                With .Columns("QTY_" & Format(i, "0")).Header
                    .Appearance.BackColor = Color.White
                    .Appearance.BackColor2 = Color.LightBlue
                    .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Column.Width = 70
                    Create_Summary(grdRSTSSPL3, "QTY_" & Format(i, "0"))
                End With
                With .Columns("AMT_" & Format(i, "0")).Header
                    .Appearance.BackColor = Color.White
                    .Appearance.BackColor2 = Color.LightGreen
                    .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Column.Width = 70
                    Create_Summary(grdRSTSSPL3, "AMT_" & Format(i, "0"))
                End With
            Next
            For Each C As String In New String() {"CUST_CODE", "ITEM_CODE", "SEASON_CODE"}
                With .Columns(C).Header
                    .Appearance.BackColor = Color.White
                    .Appearance.BackColor2 = Color.LightGray
                    .Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                .Columns(C).Width = 100
            Next
            .Columns("CUST_CODE").Header.Caption = "Customer"
            .Columns("ITEM_CODE").Header.Caption = "Item"
            .ColHeaderLines = 2
            Create_Summary(grdRSTSSPL3, "ITEM_CODE", "Count")
        End With


        grdRSTSSPLX.DataSource = dst.Tables("RSTSSPLX")
        Create_Summary(grdRSTSSPLX, "CUST_CODE", "Count")
        Create_Summary(grdRSTSSPLX, "SEL")
        Create_Summary(grdRSTSSPLX, "UPD")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTSSPLX.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Color.White
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            If New String() {"TYP_BEG", "TYP_END", "LYP_BEG", "LYP_END", "YP_INIT", "YP_LAST"}.Contains(gcol.Key) Then
                gcol.Width = 80
                gcol.Header.Appearance.TextHAlign = HAlign.Center
                gcol.CellAppearance.TextHAlign = HAlign.Center
            End If
            If New String() {"INIT_OPER", "INIT_DATE", "YP_INIT"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
            ElseIf New String() {"LAST_OPER", "LAST_DATE", "YP_LAST"}.Contains(gcol.Key) Then
                gcol.Header.Appearance.BackColor2 = Color.LightGreen
            End If
            If gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        grdRSTSSPLX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        Dim YYYY As String = Mid(ASCMAIN1.CYP, 1, 4)
        Dim NY As String = Format(Val(YYYY) + 1, "0000")
        Dim LY As String = Format(Val(YYYY) - 1, "0000")

        ASCMAIN1.sql = "Select Min (SEASON_CODE) from RSTSSPL1"
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

                If EMsg = "" Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    Dim row As DataRow = LookUp("RSTSSPL1", New String() {CUST_CODE, SEASON_CODE})
                    If row IsNot Nothing Then
                        EMsg &= vbCr & "Record Already Exists for Customer " & CUST_CODE & " in Season " & SEASON_CODE
                    End If
                End If

                If chk12Mos.Checked Then
                    EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("RSTSSPL1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Open"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                    openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME = "" Then
                    EMsg &= vbCr & "No Workbook Selected"
                Else
                    WorkbookView1.GetLock()
                    WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

                    workbook = WorkbookView1.ActiveWorkbook
                    worksheet = workbook.Worksheets(0)

                    Absx1.txtFor("CUST_CODE").Text = worksheet.Cells(0, 1).Value
                    If Absx1.cbeFor("SEASON_CODE").Value <> worksheet.Cells(3, 1).Value Then
                        EMsg &= vbCrLf & "Season must match " & Absx1.cbeFor("SEASON_CODE").Value & ", and does not (" & worksheet.Cells(3, 1).Value & ")"
                    End If
                    WorkbookView1.ReleaseLock()

                    If EMsg = "" Then
                        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                        Dim row As DataRow = LookUp("RSTSSPL1", New String() {CUST_CODE, SEASON_CODE})
                        If row Is Nothing Then
                            EMsg &= vbCr & "No Record on file for Customer " & CUST_CODE & " in Season " & SEASON_CODE
                        End If
                    End If

                    If EMsg <> "" Then
                        Validate_Code("CUST_CODE")
                        If EMsg = "" Then
                            If Not ASCMAIN1.Logical_Lock("RSTSSPL1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Edit", "View", "Update as Reqd"

                If chk12Mos.Checked Then
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                    If eItemKey <> "View" Or Not SEASON_CODE.EndsWith("S") Then
                        EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                    End If
                End If

                If EMsg = "" Then
                    blnConsolidated = False
                    If eItemKey = "View" And Absx1.txtFor("CUST_CODE").Text = "" Then
                        Dim selected_customers As Integer = dst.Tables("RSTSSPLX").Select("SEL='1'").Length
                        If MsgBox("Do you want to view the Consolidated Stock & Sales Plan for " _
                                  & IIf(selected_customers = 0, "All Customers", "the " & CStr(selected_customers) & " customers selected") & ", Combined?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            'EMsg &= vbCr & "You must first select a Customer"
                            EMsg &= vbCr & "Returning to Customer Selection"
                        Else
                            blnConsolidated = True
                            CUST_CODES_consolidated.Clear()
                            For Each row As DataRow In dst.Tables("RSTSSPLX").Select("SEL='1'", "CUST_CODE")
                                CUST_CODES_consolidated.Add(row.Item("CUST_CODE"))
                            Next
                        End If
                    Else
                        If eItemKey = "Update as Reqd" Then
                            Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                            If SEASON_CODE = "" Then
                                EMsg &= vbCr & "You must select a Season before Updating as Required'"
                            Else
                                If dst.Tables("RSTSSPLX").Select("UPD = '1'").Length = 0 Then
                                    EMsg &= vbCr & "No SSP records appear to require an update in Season " & SEASON_CODE
                                End If
                            End If

                            If EMsg = "" Then
                                If MsgBox($"OK to Update {dst.Tables("RSTSSPLX").Select("UPD = '1'").Length} Stock & Sales Plans" & vbCrLf & " Requiring Update for Season " & SEASON_CODE & "?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Exiting Update as Required"
                                End If
                            End If

                        Else
                            Validate_Code("CUST_CODE")

                            If EMsg = "" Then
                                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                                Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value & ""
                                Dim row As DataRow = LookUp("RSTSSPL1", New String() {CUST_CODE, SEASON_CODE})
                                If row Is Nothing Then
                                    EMsg &= vbCr & "No Record on file for Customer " & CUST_CODE & " in Season " & SEASON_CODE
                                End If
                            End If
                        End If
                    End If

                    If eItemKey = "Edit" Then
                        If EMsg = "" Then
                            If Not ASCMAIN1.Logical_Lock("RSTSSPL1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                                Exit Sub
                            End If
                        End If
                    End If

                End If

            Case "Actualize"

                If chk12Mos.Checked Then
                    EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                End If

                If dst.Tables("RSTSSPLX").Select("SEL = '1'", "CUST_CODE").Length = 0 Then
                    EMsg &= vbCr & "No Customers Selected"
                Else
                    If EMsg = "" Then
                        Dim CUST_CODEs As New List(Of String)
                        For Each row As DataRow In dst.Tables("RSTSSPLX").Select("SEL = '1'", "CUST_CODE")
                            Dim CUST_CODE As String = row.Item("CUST_CODE")
                            If Not ASCMAIN1.Logical_Lock("RSTSSPL1", "CUST_CODE" & ":" & CUST_CODE) Then
                                Exit Sub
                            End If
                        Next
                    End If
                End If


            Case "Publish Budget"

                If chk12Mos.Checked Then
                    EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                End If
                If EMsg = "" Then
                    If MsgBox("This command will Publish All Stock & Sales Plans for " & Absx1.cbeFor("SEASON_CODE").Value _
                          & vbCrLf & " and then tie the total for the Channel into the Financial Budget for Sales" _
                          & vbCrLf & " by making an Adjusting Entry using the IPLBFINx customers." _
                          & vbCrLf & vbCrLf & "If you Continue, you will first see a grid comparing" _
                          & vbCrLf & " the Financial Budget to the Sales Budget, showing the Adjusting Entry," _
                          & vbCrLf & " and then you will need to click Update to actually Publish the Budget." _
                          & vbCrLf & vbCrLf & "OK to Continue?", MsgBoxStyle.OkCancel, "Verification") = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                If EntryMode = "P" Then
                ElseIf EntryMode = "I" Then

                    CUST_CODEs_import = ""
                    For Each row As DataRow In ASCDATA1.SelectDistinct("RSTSSPL3", "CUST_CODE").Select("")
                        Dim CUST_CODE As String = row.Item(0)
                        If LookUp("ARTCUST1", CUST_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                        End If
                        CUST_CODEs_import &= "," & CUST_CODE
                    Next

                    ITEM_CODEs_import = ""
                    For Each row As DataRow In ASCDATA1.SelectDistinct("RSTSSPL3", "ITEM_CODE").Select("")
                        Dim ITEM_CODE As String = row.Item(0)
                        If LookUp("ICTITEM1", ITEM_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Item Code " & ITEM_CODE
                        End If
                        ITEM_CODEs_import &= "," & ITEM_CODE
                    Next

                    If EMsg = "" Then
                        If MsgBox("OK to Update the Stock and Sales Plans for the following customers in " & Absx1.cbeFor("SEASON_CODE").Value & "?" _
                            & vbCrLf & Mid(CUST_CODEs_import, 2), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Else
                            Exit Sub
                        End If
                    End If
                Else
                    WorkbookView1.EndEdit()

                    Process_Workbook()

                End If

            Case "Load SSP Forecasts"
                If chk12Mos.Checked Then
                    EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                End If

                If EMsg = "" Then

                    If MsgBox("This function will Import Unit and Sales Forecast data" _
                & vbCrLf & " from a specifically formatted spreadsheet" _
                & vbCrLf & " and use that data to replace Stock & Sales Plans currently on file" _
                & vbCrLf & " in Season " & Absx1.cbeFor("SEASON_CODE").Value _
                & vbCrLf & " for all of the Customers represented in the imported spreadsheet" _
                & vbCrLf _
                & vbCrLf & "Once you click 'Yes' to proceed," _
                & vbCrLf & " you will be asked for the location of the spreadsheet, " _
                & vbCrLf & " and the data will be imported and displayed in a grid." _
                & vbCrLf _
                & vbCrLf & "You will have an opportunity to review it before clicking 'Update'." _
                & vbCrLf _
                & vbCrLf & "Proceed with the Import?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Customer HC Pivot"

                If chk12Mos.Checked Then
                    EMsg &= vbCr & "12 Mos option is available for View Only, and only when selecting a Spring season"
                End If

                If EMsg = "" Then

                    Dim selected_customers As Integer = dst.Tables("RSTSSPLX").Select("SEL='1'").Length
                    If selected_customers = 0 Then
                        EMsg &= vbCr & "You must select at least 1 customer"
                    Else
                        '

                        If MsgBox("Do you want to view the Customer High Collection Pivot for " _
                              & IIf(selected_customers = 0, "All Customers", "the " & CStr(selected_customers) & " customers selected") & "?",
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Returning to Customer Selection"
                        Else
                            blnConsolidated = True
                            CUST_CODES_consolidated.Clear()
                            For Each row As DataRow In dst.Tables("RSTSSPLX").Select("SEL='1'", "CUST_CODE")
                                CUST_CODES_consolidated.Add(row.Item("CUST_CODE"))
                            Next
                        End If
                    End If
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

                If blnConsolidated Then
                    Prepare_RollUp()
                End If

            Case "Open"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "P" Then
                    Update_Record_Publish()
                ElseIf EntryMode = "I" Then
                    Update_Record_Import()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Update as Reqd"

                updateAsRequired = True

                Dim ctr As Integer = 0
                Dim CUST_CODEs As New List(Of String)
                For Each row As DataRow In dst.Tables("RSTSSPLX").Select("UPD = '1'")
                    ctr += 1
                    CUST_CODEs.Add(row.Item("CUST_CODE"))
                    'If ctr = 2 Then Exit For ' for testing
                Next

                For Each CUST_CODE As String In CUST_CODEs
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                    Click_Command("Edit")
                    Click_Command("Update")
                Next

                MsgBox("All Reqired Updates have been Completed", MsgBoxStyle.OkOnly)
                updateAsRequired = False
                Mode_Settings(False)

            Case "Cancel", "Done"
                blnConsolidated = False
                Mode_Settings(False)

            Case "Save XLSX"
                WorkbookView1.GetLock()
                Dim FILENAME_PFX As String = ""
                If automated_XLS Then
                    FILENAME_PFX = SEASON_CODE & "_" & CUST_CODE & "_"
                End If
                Dim FILENAME As String = ASCMAIN1.Folders("Work") & FILENAME_PFX & ASCMAIN1.Next_Control_No("RSFSSPL1.XLSX_NO") & ".XLSX"

                If filename_special <> "" Then ' running in Actualize Mode
                    FILENAME = filename_special
                End If


                WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

                If filename_special <> "" Or automated_XLS Then ' running in Actualize Mode
                    ' no show
                Else
                    Show_Document(FILENAME)
                End If

                WorkbookView1.ReleaseLock()

            Case "Publish Budget"
                EntryMode = "P"
                Load_Record()
                Mode_Settings(True)

            Case "Load SSP Forecasts"

                If Excel_Import_SG_from_Forecast_File() = -1 Then ' Excel_Import_SG(grdRSTSSPL3) = -1 Then
                    ' nothing happened
                Else
                    Sort_grdColumns(grdRSTSSPL3, "CUST_CODE,ITEM_CODE")
                    'Setup_grd()
                    Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value
                    Dim SEASON_YEAR As String = Mid(SEASON_CODE, 1, 4)
                    Dim SEASON_TYPE As String = Mid(SEASON_CODE, 5, 1)
                    grdRSTSSPL3.Text = "Sales Forecasts by Customer, Item, Month for Season " & Absx1.cbeFor("SEASON_CODE").Value

                    YPs_Imported.Clear()

                    With grdRSTSSPL3.DisplayLayout.Bands(0)
                        For i As Integer = 1 To 6
                            Dim YP As String = SEASON_YEAR & Format(i + IIf(SEASON_TYPE = "F", 6, 0), "00")
                            YPs_Imported.Add(YP)
                            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                            Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                            With .Columns("QTY_" & Format(i, "0")).Header
                                .Caption = Mid(LEGEND, 10, 6) & vbCrLf & "Qty"
                            End With
                            With .Columns("AMT_" & Format(i, "0")).Header
                                .Caption = Mid(LEGEND, 10, 6) & vbCrLf & "Amt"
                            End With
                        Next
                    End With

                    EntryMode = "I"
                    ' Load_Record()
                    Mode_Settings(True)
                End If

            Case "Actualize"

                Dim CUST_CODEs As New List(Of String)
                For Each row As DataRow In dst.Tables("RSTSSPLX").Select("SEL = '1'", "CUST_CODE")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    CUST_CODEs.Add(CUST_CODE)
                Next

                Dim XNO As String = ASCMAIN1.Next_Control_No("RSFSSPL1.ACTUALIZE")
                'Dim FOLDER As String = "S:\INT\SSP\" & XNO & "\"
                Dim FOLDER As String = ASCMAIN1.Folders("SharedRoot") & "\SSP\" & XNO & "\"
                Dim YYYYMMDD As String = Format(Now, "yyyyMMdd")

                My.Computer.FileSystem.CreateDirectory(FOLDER)
                For Each CUST_CODE As String In CUST_CODEs
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                    Click_Command("Edit")
                    'filename_special = FOLDER & CUST_CODE & ".XLSX"

                    filename_special = FOLDER & CUST_CODE & "_" & YYYYMMDD & ".XLSX"
                    Click_Command("Save XLSX")
                    Click_Command("Update")
                    filename_special = ""
                Next

                Process.Start(FOLDER)

            Case "Customer HC Pivot"

                Customer_HC_Pivot()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Open").Settings.Enabled = not_iScreenMode
                    .Items("Save XLSX").Visible = ScreenMode And (EntryMode <> "P") And (EntryMode <> "I")

                    .Items("Publish Budget").Visible = Not ScreenMode And (ASCMAIN1.Running_in_VS Or Split(ASCMAIN1.USER_SECURITY_CODEs, ",").Contains("PB"))
                    .Items("Actualize").Visible = Not ScreenMode And (ASCMAIN1.Running_in_VS Or Split(ASCMAIN1.USER_SECURITY_CODEs, ",").Contains("PB"))
                    .Items("Load SSP Forecasts").Visible = Not ScreenMode

                    If EntryMode = "V" And ScreenMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        .Items("Update as Reqd").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                        .Items("Update as Reqd").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L" Or EntryMode = "P" Or EntryMode = "I")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L" Or EntryMode = "P" Or EntryMode = "I")

                    .Items("New").Visible = Not .Items("Update").Visible And Not blnConsolidated
                    .Items("Edit").Visible = Not .Items("Update").Visible And Not blnConsolidated
                    .Items("Update as Reqd").Visible = Not .Items("Update").Visible And Not blnConsolidated
                    .Items("Open").Visible = Not .Items("Update").Visible And Not blnConsolidated

                    .Items("Open").Visible = False

                    .Items("Done").Visible = ScreenMode And EntryMode = "V"

                    .Items("Customer HC Pivot").Visible = Not ScreenMode

                End With

                .Groups("Display Options").Visible = ScreenMode And (EntryMode <> "P") And (EntryMode <> "I") And Not blnConsolidated
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdRSTSSPLX.Visible = Not ScreenMode
        splSSPL.Visible = ScreenMode And EntryMode <> "P" And EntryMode <> "I"
        grdSATBUDWX.Visible = ScreenMode And EntryMode = "P"
        grdRSTSSPL3.Visible = ScreenMode And EntryMode = "I"
        spl.Panel1Collapsed = (EntryMode = "P")

        If ScreenMode Then
            If EntryMode = "P" Then
                Dim SD As Date = CDate(IIf(SEASON_TYPE = "S", "01", "07") & "/01/" & SEASON_YEAR)
                With grdSATBUDWX.DisplayLayout.Bands(0)
                    For M As Integer = 1 To 6
                        Dim D As Date = SD.AddMonths(M - 1)
                        Dim LEGEND As String = Format(D, "MMM")
                        .Columns("RBUD_" & Format(M, "0")).Header.Caption = "Rtl " & LEGEND
                        .Columns("PBUD_" & Format(M, "0")).Header.Caption = "Old " & LEGEND
                        .Columns("SBUD_" & Format(M, "0")).Header.Caption = "New " & LEGEND
                        .Columns("FBUD_" & Format(M, "0")).Header.Caption = "Fin " & LEGEND
                        .Columns("XBUD_" & Format(M, "0")).Header.Caption = "Adj " & LEGEND
                    Next
                    .Columns("RBUD_0").Header.Caption = "Rtl Sls"
                    .Columns("PBUD_0").Header.Caption = "Old Sls"
                    .Columns("SBUD_0").Header.Caption = "New Sls"
                    .Columns("FBUD_0").Header.Caption = "Finance"
                    .Columns("XBUD_0").Header.Caption = "Adjust"
                End With

            End If
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "RSTSSPL1", "RSTSSPL2", "SATBUDWX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        SEASON_CODE = ""
        chkShowLINE_TAGs.Checked = False
        blnConsolidated = False

        Absx1.txtFor("CUST_CODE").Text = ""

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

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        SEASON_TYPE = rowICTSEAS1.Item("SEASON_TYPE")
        SEASON_YEAR = rowICTSEAS1.Item("SEASON_YEAR")
        DTE1 = CDate(IIf(SEASON_TYPE = "S", "02", "08") & "/01/" & SEASON_YEAR)
        DTE2 = DTE1.AddMonths(6).AddDays(-1)

        SEASON_CODE_prior = IIf(SEASON_TYPE = "S",
                                Format(Val(SEASON_YEAR) - 1, "0000") & "F",
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


        If EntryMode = "P" Then
            Prepare_Summary_Budget()

        Else

            If EntryMode = "N" Or blnConsolidated Then
                rowRSTSSPL1 = dst.Tables("RSTSSPL1").NewRow
                rowRSTSSPL1.Item("CUST_CODE") = CUST_CODE
                rowRSTSSPL1.Item("SEASON_CODE") = SEASON_CODE
                rowRSTSSPL1.Item("INIT_DATE") = DATETIME_STAMP
                rowRSTSSPL1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowRSTSSPL1.Item("YP_INIT") = ASCMAIN1.CYP

                XLS_NO = ASCMAIN1.Next_Control_No("RSTSSPL1.XLS_NO")
                rowRSTSSPL1.Item("XLS_NO") = XLS_NO

                dst.Tables("RSTSSPL1").Rows.Add(rowRSTSSPL1)
            Else
                rowRSTSSPL1 = Fill_Record("RSTSSPL1", New String() {CUST_CODE, SEASON_CODE})
                rowRSTSSPL1.Item("YP_LAST") = ASCMAIN1.CYP
                XLS_NO = rowRSTSSPL1.Item("XLS_NO") & ""
            End If

            EnforceConstraints(False)

            If Not blnConsolidated Then
                rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
                rowSOTTCLS1 = LookUp("SOTTCLS1", rowARTCUST1.Item("TRADE_CLASS_CODE"))
                rowSOTPCLS1 = LookUp("SOTPCLS1", rowARTCUST1.Item("PRICE_CLASS_CODE"))
            End If

            EnforceConstraints(True)

            If EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V" Then
                Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
                If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.CLIENT = "INT" Then 'And Format(Now, "MM/dd/yyyy") = "01/13/2025" Then
                    Stop
                    FILENAME = "C:\Share\INT\Templates\" & Me.Name & ".xlsx"
                    'FILENAME = "C:\Users\nicholas\Desktop\RSFSSPL1_Copy.xlsx"
                End If
                WorkbookView1.GetLock()
                WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                XLS_Validation(True)

                XLS_Refresh_HC()
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

                    Dim xSEASON_CODE As String = SEASON_CODE
                    Dim xSEASON_CODE_prior As String = SEASON_CODE_prior

                    For ST As Integer = 0 To If(chk12Mos.Checked, 1, 0)
                        Dim STA As Integer = If(ST = 1, 6, 0)
                        If ST = 1 Then
                            mid(xSEASON_CODE, 5, 1) = "F"
                            mid(xSEASON_CODE_prior, 5, 1) = "S"
                        End If

                        If blnConsolidated Then
                            ASCMAIN1.sql = "Select '*' CUST_CODE, SEASON_CODE, HC_CODE, LINE_TAG, LINE_KEY" & vbCrLf _
                                        & ", SUM (AMT_0) AMT_0" & vbCrLf _
                                        & ", SUM (AMT_1) AMT_1, SUM (AMT_2) AMT_2, SUM (AMT_3) AMT_3" & vbCrLf _
                                        & ", SUM (AMT_4) AMT_4, SUM (AMT_5) AMT_5, SUM (AMT_6) AMT_6" & vbCrLf _
                                        & ", 0 AMT_X, NULL NOTES" & vbCrLf _
                                        & " from RSTSSPL2 where SEASON_CODE = '" & xSEASON_CODE & "'" & vbCrLf _
                                        & IIf(CUST_CODES_consolidated.Count = 0, "", " and CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')" & vbCrLf) _
                                        & " group by SEASON_CODE, HC_CODE, LINE_TAG, LINE_KEY"

                            Fill_Records("RSTSSPL2", "", True, ASCMAIN1.sql)

                        Else
                            Fill_Records("RSTSSPL2", New String() {CUST_CODE, xSEASON_CODE})

                        End If

                        Dim HC_CODE As String = ""
                        Dim HC_CODEs As New List(Of String)
                        For Each row As DataRow In dst.Tables("RSTSSPL2").Select("", "HC_CODE")
                            If row.Item("HC_CODE") <> HC_CODE Then
                                HC_CODE = row.Item("HC_CODE")
                                worksheet = workbook.Worksheets(HC_CODE)
                            End If

                            If worksheet IsNot Nothing Then
                                If Not HC_CODEs.Contains(HC_CODE) Then HC_CODEs.Add(HC_CODE)

                                'Debug.Print(worksheet.Cells("R7").Formula)

                                Dim LINE_TAG As String = row.Item("LINE_TAG")
                                If LINE_TAG = "PSIITM" Or LINE_TAG = "TOAITM" Then
                                    Dim LINE_KEY As String = row.Item("LINE_KEY")
                                    If XLS_Allocation_Lines.Keys.Contains(LINE_KEY) Then
                                        Dim r As Integer = XLS_Allocation_Lines(LINE_KEY)
                                        For m As Integer = ST To 6

                                            If ASCMAIN1.Running_in_VS AndAlso worksheet.Cells(r, c0 + m + STA).Formula & "" <> "" Then Stop ' FOR RSTSSPP2 BELOW

                                            If worksheet.Cells(r, c0 + m + STA).Formula & "" = "" Then
                                                worksheet.Cells(r, c0 + m + STA).Value = row.Item("AMT_" & CStr(m))
                                            End If
                                        Next
                                    End If

                                    If LINE_TAG = "PSIITM" Then
                                        Dim SQLW As String = $"CUST_CODE = '{CUST_CODE}' and HC_CODE = '{HC_CODE}' and ALLO_CTL_NO = '{LINE_KEY}'"
                                        Dim rows() As DataRow = dst.Tables("RSTSSPP2").Select(SQLW)
                                        If rows.Length <> 0 Then
                                            For m As Integer = ST To 6
                                                rows(0).Item("M0" & Format(m, "0")) = row.Item("AMT_" & CStr(m))
                                            Next
                                        End If
                                    End If
                                Else

                                    If LINE_TAG = "TYPSLS%LY" Or LINE_TAG = "TYPSHP%LY" Then
                                    Else
                                        For m As Integer = ST To 6
                                            If worksheet.Cells(XLSR(LINE_TAG), c0 + m + STA).Formula & "" = "" Then
                                                worksheet.Cells(XLSR(LINE_TAG), c0 + m + STA).Value = row.Item("AMT_" & CStr(m))
                                            End If
                                        Next
                                    End If

                                    If New String() {"TYPBOM", "LYABOM"}.Contains(LINE_TAG) Then

                                        'worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + STA + 3).Value = row.Item("AMT_X") 'WIPES OUT OCTOBER HERE, DOOR COUNT SHOWS AT 0 FOR EVERY SINGLE HC
                                        If chk12Mos.Checked Then
                                            worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + 6 + 3).Value = row.Item("AMT_X")
                                        Else
                                            worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + STA + 3).Value = row.Item("AMT_X")
                                        End If

                                    End If

                                    If New String() {"TYPBOM", "TYPSLS", "TYPSLSACT", "TYPSLSACTZ", "TYPGRSB", "TYPGRSP", "TYPGRS", "TYPADJ", "TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET", "TYPEOM", "TYPSHP%LY", "TYPSLS%LY"}.Contains(LINE_TAG) Then

                                        Dim offset As Integer = 0
                                        If chk12Mos.Checked Then offset = 6

                                        worksheet.Cells(XLSR(LINE_TAG), c0 + 6 + offset + 3 + 4).Value = row.Item("NOTES")

                                    End If

                                    If blnConsolidated And New String() {"TYPGRSBNET", "TYPGRSPNET", "TYPGRSNET"}.Contains(LINE_TAG) Then
                                        For m As Integer = 0 To 6
                                            ' worksheet.Cells(XLSR(LINE_TAG), c0 + m + STA).Value = row.Item("AMT_" & CStr(m))
                                            worksheet.Cells(XLSR(LINE_TAG), c0 + m + STA).Value = row.Item("AMT_" & CStr(m))
                                        Next
                                    End If
                                End If
                            End If
                        Next

                        If automated_Customer_HC_Pivot Then

                            For Each HC_CODE In HC_CODEs
                                worksheet = workbook.Worksheets(HC_CODE)

                                Dim LINE_NO_ctr As Integer = 0
                                For Each LT As String In New String() {"TYPSLSACTZ", "TYPGRSB", "TYPGRSP", "TYPEOM"}
                                    Dim rowRSTSSPP1 As DataRow = dst.Tables("RSTSSPP1").NewRow
                                    With rowRSTSSPP1
                                        .Item("SEASON_CODE") = xSEASON_CODE
                                        .Item("CUST_CODE") = CUST_CODE
                                        .Item("HC_CODE") = HC_CODE
                                        LINE_NO_ctr += 1
                                        .Item("LINE_NO") = LINE_NO_ctr
                                        .Item("LINE_DESC") = worksheet.Cells(XLSR(LT), 3).Value
                                        For M As Integer = 0 To 6
                                            .Item("M0" & Format(M, "0")) = Val(worksheet.Cells(XLSR(LT), 4 + M).Value & "")
                                        Next
                                    End With
                                    dst.Tables("RSTSSPP1").Rows.Add(rowRSTSSPP1)
                                Next
                            Next

                        End If


                        If EntryMode = "V" Then
                            ASCMAIN1.sql = "Select HC_CODE, LINE_TAG" & vbCrLf _
                            & ", SUM (NVL(AMT_0,0)+NVL(AMT_1,0)+NVL(AMT_2,0)+NVL(AMT_3,0)+NVL(AMT_4,0)+NVL(AMT_5,0)) AMT" & vbCrLf _
                            & " from RSTSSPL2 " & vbCrLf _
                            & " where SEASON_CODE = '" & xSEASON_CODE_prior & "'" & vbCrLf _
                            & IIf(CUST_CODES_consolidated.Count = 0, "", " and CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')" & vbCrLf) _
                            & $"   and LINE_TAG IN ({LINE_TAGs_Roll_Up})" & vbCrLf _
                            & " group by HC_CODE, LINE_TAG"
                            Fill_Records("RSTSSPLR", "", True, ASCMAIN1.sql)
                        End If
                    Next

                End If

                WorkbookView1.ReleaseLock()

                'If EntryMode = "N" Then
                Get_LY_Actuals()
                'End If
            Else
                WorkbookView1.GetLock()

                XLS_Validation(False)

                'XLS_Refresh_HC()
                'XLS_Refresh_Allocations()

                ' xLOAD WORKBOOK FROM DISK
                ' xREFRESH CERTAIN LINES
                ' ADD NEW ALLOCATION ITEMS
                ' REFRESH ALLOCATION QTYS
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
            Get_Retailer_BOM_EOW()

            If chk12Mos.Checked AndAlso XLSR.ContainsKey("RTLOHB") Then
                WorkbookView1.GetLock()
                Try
                    For Each ws As SpreadsheetGear.IWorksheet In workbook.Worksheets
                        If ws.Index > 0 AndAlso
                           Not ws.Name.Equals("Roll-Up", StringComparison.OrdinalIgnoreCase) Then
                            Recalc_RTLOHB_Forward(ws, c0 + 7)
                        End If
                    Next
                    Dim extraMonths As Integer = 6
                    Dim lastMonthCol As Integer = c0 + 6 + extraMonths
                    For c As Integer = c0 + 1 To lastMonthCol
                        WriteTotalRTLOHBFormula(c)
                    Next

                    workbook.WorkbookSet.Calculate()
                Finally
                    WorkbookView1.ReleaseLock()
                End Try
            End If
            WorkbookView1.GetLock()
            For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
                If worksheet.ProtectContents Then worksheet.Unprotect(XLS_PWD)
                'worksheet.Cells(XLSR("TYPGRSBNET"), 0).EntireRow.Hidden = True
                'worksheet.Cells(XLSR("TYPGRSPNET"), 0).EntireRow.Hidden = True

                worksheet.Protect(XLS_PWD)
            Next
            WorkbookView1.ReleaseLock()
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Dim sqld As String = "CUST_CODE = '" & CUST_CODE & "' and SEASON_CODE = '" & SEASON_CODE & "'"
        INIT_LAST("RSTSSPL1")
        Update_Record_TDA("RSTSSPL1", sqld)
        Update_Record_TDA("RSTSSPL2", sqld)

        If filename_special <> "" Or updateAsRequired Then ' running in Actualize Mode
            CommitTrans("")
        Else
            CommitTrans("Update Complete")
        End If

    End Sub

    Sub Process_Workbook()

        dst.Tables("RSTSSPL2").Rows.Clear()

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
        'Dim sqld As String = "SATBUDW1.OPS_YYYY = '" & SEASON_YEAR & "'" & vbCrLf _
        '                     & " and (SATBUDW1.CUST_CODE = 'IPLBFIN' or SATBUDW1.CUST_CODE in " & vbCrLf _
        '                     & "(Select ARTCUST1.CUST_CODE from ARTCUST1,SOTTCLS1" & vbCrLf _
        '                     & " where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
        '                     & "   and SOTTCLS1.AUTH_REQD = '1'))"

        Dim sqld As String = "SATBUDW1.OPS_YYYY = '" & SEASON_YEAR & "'"
        Return sqld
    End Function

    Sub Update_Record_Import()

        Dim CUST_CODEs_not_MARKETS As New List(Of String)
        dst.Tables("DPTITMF1").Rows.Clear()
        Dim YPs_Importedx As String = "'" & Join(YPs_Imported.ToArray, "','") & "'"
        If YPs_Imported.Contains(ASCMAIN1.CYP) Then
            YPs_Importedx = "'000000'," & YPs_Importedx
        End If

        BeginTrans()

        Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value

        Dim sqld As String = "SEASON_CODE = '" & SEASON_CODE & "'" _
                             & " and CUST_CODE in ('" & Replace(CUST_CODEs_import, ",", "','") & "')"
        Update_Record_TDA("RSTSSPL3", sqld)


        ASCDATA1.ExecuteSQL("Delete from RSTSSPL1 where " & sqld)
        ASCDATA1.ExecuteSQL("Delete from RSTSSPL2 where " & sqld)

        dst.Tables("RSTSSPL1").Rows.Clear()
        dst.Tables("RSTSSPL2").Rows.Clear()
        For Each CUST_CODE As String In Split(Mid(CUST_CODEs_import, 2), ",")
            Dim rowRSTSSPL1 As DataRow = dst.Tables("RSTSSPL1").NewRow
            With rowRSTSSPL1
                .Item("CUST_CODE") = CUST_CODE
                .Item("SEASON_CODE") = SEASON_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("SEASON_CODE") = SEASON_CODE
            End With
            dst.Tables("RSTSSPL1").Rows.Add(rowRSTSSPL1)

            Dim sql As String = ""
            For i As Integer = 1 To 6
                Dim C As String = "AMT_" & Format(i, "0")
                sql &= ", Sum (RSTSSPL3." & C & ") " & C & vbCrLf
            Next
            ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
                & ", DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','TYPGRSBNET','TYPGRSPNET') LINE_TAG" & vbCrLf _
                & sql _
                & " from RSTSSPL3,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = RSTSSPL3.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and RSTSSPL3.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and RSTSSPL3.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & " group by ICTCOLL1.HC_CODE, DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','TYPGRSBNET','TYPGRSPNET')"

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim HC_CODE As String = row.Item("HC_CODE")
                Dim LINE_TAG As String = row.Item("LINE_TAG")
                Dim rowRSTSSPL2 As DataRow = dst.Tables("RSTSSPL2").NewRow
                With rowRSTSSPL2
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("SEASON_CODE") = SEASON_CODE
                    .Item("HC_CODE") = HC_CODE
                    .Item("LINE_TAG") = LINE_TAG
                    .Item("LINE_KEY") = "X"
                    For I As Integer = 1 To 6
                        Dim C As String = "AMT_" & Format(I, "0")
                        Dim C2 As String = "AMT_" & Format(I - 1, "0")
                        .Item(C2) = row.Item(C)
                    Next
                End With
                dst.Tables("RSTSSPL2").Rows.Add(rowRSTSSPL2)
            Next

            ' Get Forecast record for this customer

            ASCMAIN1.sql = "Select * from SOTMKTC1 where CUST_CODE = :PARM1"
            Dim rowSOTMKTC1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", CUST_CODE)
            If rowSOTMKTC1 Is Nothing Then
                CUST_CODEs_not_MARKETS.Add(CUST_CODE)
            Else
                Dim MARKET_CODE As String = rowSOTMKTC1.Item("MARKET_CODE")
                Dim sql_Delete As String = "Delete from DPTITMF1" _
                    & " where MARKET_CODE = '" & MARKET_CODE & "'" _
                    & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
                    & "   and OPS_YYYYPP_FC in (" & YPs_Importedx & ")"
                ASCDATA1.ExecuteSQL(sql_Delete)

                For Each rowRSTSSPL3 As DataRow In dst.Tables("RSTSSPL3").Select("CUST_CODE = '" & CUST_CODE & "'")
                    For P As Integer = 1 To 6
                        Dim FORECAST As Integer = Val(rowRSTSSPL3.Item("QTY_" & Format(P, "0")) & "")
                        Dim OPS_YYYYPP_FC As String = YPs_Imported(P - 1)
                        If FORECAST <> 0 And OPS_YYYYPP_FC >= ASCMAIN1.CYP Then
                            Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                            rowDPTITMF1.Item("MARKET_CODE") = MARKET_CODE
                            rowDPTITMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                            rowDPTITMF1.Item("ITEM_CODE") = rowRSTSSPL3.Item("ITEM_CODE")

                            rowDPTITMF1.Item("OPS_YYYYPP_FC") = OPS_YYYYPP_FC
                            rowDPTITMF1.Item("FORECAST") = FORECAST
                            dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
                        End If
                    Next
                Next
            End If
        Next

        Update_Record_TDA("DPTITMF1")

        Update_Record_TDA("RSTSSPL1")
        Update_Record_TDA("RSTSSPL2")

        If CUST_CODEs_not_MARKETS.Count <> 0 Then
            MsgBox("Please Note - the following customers are not linked to a market," _
                   & vbCrLf & " so no forecast changes were made for these customers:" _
                   & vbCrLf & Join(CUST_CODEs_not_MARKETS.ToArray, ","))
        End If

        CommitTrans("Update Complete")

    End Sub

    Sub Update_Record_Publish()
        Dim sqld As String = Get_SQL_SATBUDW1()

        Dim O As Integer = IIf(SEASON_TYPE = "S", 0, 6)
        For Each rowSATBUDW1 As DataRow In dst.Tables("SATBUDW1").Select("")
            For I As Integer = 1 To 6
                rowSATBUDW1.Item("WB_P" & Format(I + O, "00")) = 0
            Next
        Next

        'ASCMAIN1.sql = "Select * from " & RSTSSPL2 & " where NVL(AMT_1,0) <> 0 or NVL(AMT_2,0) <> 0 or NVL(AMT_3,0) <> 0 or NVL(AMT_4,0) <> 0 or NVL(AMT_5,0) <> 0 or NVL(AMT_6,0) <> 0"
        ASCMAIN1.sql = "Select * from " & RSTSSPL2 & " where NVL(AMT_0,0) <> 0 or NVL(AMT_1,0) <> 0 or NVL(AMT_2,0) <> 0 or NVL(AMT_3,0) <> 0 or NVL(AMT_4,0) <> 0 or NVL(AMT_5,0) <> 0"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim OPS_YYYY As String = SEASON_YEAR
            Dim COLLECTION_CODE As String = ROW.Item("COLLECTION_CODE")
            Dim ITEM_BASIC_PROMO As String = ROW.Item("ITEM_BASIC_PROMO")
            Dim CUST_CODE As String = ROW.Item("CUST_CODE")
            Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find _
                    (New String() {OPS_YYYY, COLLECTION_CODE, ITEM_BASIC_PROMO, CUST_CODE})
            If rowSATBUDW1 Is Nothing Then
                rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                rowSATBUDW1.Item("OPS_YYYY") = OPS_YYYY
                rowSATBUDW1.Item("COLLECTION_CODE") = COLLECTION_CODE
                rowSATBUDW1.Item("ITEM_BASIC_PROMO") = ITEM_BASIC_PROMO
                rowSATBUDW1.Item("CUST_CODE") = CUST_CODE
                dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
            End If
            For I As Integer = 1 To 6
                rowSATBUDW1.Item("WB_P" & Format(I + O, "00")) = Val(rowSATBUDW1.Item("WB_P" & Format(I + O, "00")) & "") + Val(ROW.Item("AMT_" & CStr(I - 1)) & "") * 1000
            Next
        Next

        For Each rowC As DataRow In ASCDATA1.SelectDistinct("SATBUDWX", New String() {"CHANNEL_CODE"}).Select("", "CHANNEL_CODE")
            Dim CHANNEL_CODE As String = rowC.Item("CHANNEL_CODE")
            Dim rowSOTCHAN1 As DataRow = LookUp("SOTCHAN1", CHANNEL_CODE)
            Dim CUST_CODE As String = rowSOTCHAN1.Item("CUST_CODE")

            For Each ROW As DataRow In dst.Tables("SATBUDWX").Select("CHANNEL_CODE= '" & CHANNEL_CODE & "' and (ISNULL(XBUD_1,0) <> 0 or ISNULL(XBUD_2,0) <> 0 or ISNULL(XBUD_3,0) <> 0 or ISNULL(XBUD_4,0) <> 0 or ISNULL(XBUD_5,0) <> 0 or ISNULL(XBUD_6,0) <> 0)")
                Dim OPS_YYYY As String = SEASON_YEAR
                Dim HC_CODE As String = ROW.Item("HC_CODE")
                Dim rowICTCOLL0 As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
                Dim COLLECTION_CODE As String = rowICTCOLL0.Item("COLLECTION_CODE")
                Dim ITEM_BASIC_PROMO As String = "B"
                Dim rowSATBUDW1 As DataRow = dst.Tables("SATBUDW1").Rows.Find _
                        (New String() {OPS_YYYY, COLLECTION_CODE, ITEM_BASIC_PROMO, CUST_CODE})
                If rowSATBUDW1 Is Nothing Then
                    rowSATBUDW1 = dst.Tables("SATBUDW1").NewRow
                    rowSATBUDW1.Item("OPS_YYYY") = OPS_YYYY
                    rowSATBUDW1.Item("COLLECTION_CODE") = COLLECTION_CODE
                    rowSATBUDW1.Item("ITEM_BASIC_PROMO") = ITEM_BASIC_PROMO
                    rowSATBUDW1.Item("CUST_CODE") = CUST_CODE
                    dst.Tables("SATBUDW1").Rows.Add(rowSATBUDW1)
                End If
                For I As Integer = 1 To 6
                    rowSATBUDW1.Item("WB_P" & Format(I + O, "00")) = Val(ROW.Item("XBUD_" & CStr(I)) & "")
                Next
            Next
        Next

        ASCDATA1.DeleteRows("SATBUDW1", "WB_P00 = 0")

        BeginTrans()

        Update_Record_TDA("SATBUDW1", sqld)

        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.sql = "SELECT SUM (BUDGET_TYTM) FROM (
        Select SOTTCLS1.CHANNEL_CODE AS CHANNEL_CODE, SATBUDW1.CUST_CODE AS CUST_CODE
        , SUM (NVL(SATBUDW1.WB_P07,0)/1000) BUDGET_TYTM
        , SUM (NVL(SATBUDW1.WB_P01,0)/1000+NVL(SATBUDW1.WB_P02,0)/1000+NVL(SATBUDW1.WB_P03,0)/1000+NVL(SATBUDW1.WB_P04,0)/1000+NVL(SATBUDW1.WB_P05,0)/1000+NVL(SATBUDW1.WB_P06,0)/1000+NVL(SATBUDW1.WB_P07,0)/1000+NVL(SATBUDW1.WB_P08,0)/1000) BUDGET_TYTD
        , SUM (NVL(SATBUDW1.WB_P01,0)/1000+NVL(SATBUDW1.WB_P02,0)/1000+NVL(SATBUDW1.WB_P03,0)/1000+NVL(SATBUDW1.WB_P04,0)/1000+NVL(SATBUDW1.WB_P05,0)/1000+NVL(SATBUDW1.WB_P06,0)/1000+NVL(SATBUDW1.WB_P07,0)/1000+NVL(SATBUDW1.WB_P08,0)/1000+NVL(SATBUDW1.WB_P09,0)/1000+NVL(SATBUDW1.WB_P10,0)/1000+NVL(SATBUDW1.WB_P11,0)/1000+NVL(SATBUDW1.WB_P12,0)/1000) BUDGET_TYTOT
         from SATBUDW1 SATBUDW1,SOTTCLS1,ARTCUST1
         where  SOTTCLS1.CHANNEL_CODE IN ('1','2','4') AND ARTCUST1.CUST_CODE = SATBUDW1.CUST_CODE AND SOTTCLS1.TRADE_CLASS_CODE(+) = ARTCUST1.TRADE_CLASS_CODE 
        and SATBUDW1.OPS_YYYY = '2025'
         group by SOTTCLS1.CHANNEL_CODE, SATBUDW1.CUST_CODE
        ) WHERE CHANNEL_CODE = '1'"
            Dim BUDGET As Decimal = Val(ASCDATA1.GetDataValue)
            Debug.Print(BUDGET) ' TO TEST BUDGET
            Stop
        End If

        CommitTrans("Update Complete")

    End Sub
    Sub Save_LINE_TAG(LINE_TAG As String,
                      ws As SpreadsheetGear.IWorksheet,
                      Optional XLR As Integer = -1)

        Dim LINE_KEY As String = "X"
        If XLR = -1 Then
            XLR = XLSR(LINE_TAG)
        Else
            LINE_KEY = ws.Cells(XLR, 1).Value
        End If

        Dim row As DataRow = dst.Tables("RSTSSPL2").NewRow
        row.Item("CUST_CODE") = CUST_CODE
        row.Item("SEASON_CODE") = SEASON_CODE
        row.Item("HC_CODE") = ws.Name
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

        dst.Tables("RSTSSPL2").Rows.Add(row)
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
        Load_Popup_Menu(grdRSTSSPLX, "SSBBBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Select All for Channel", "Export All to XLS")
        Load_Popup_Menu(grdRSTSSPL3, "SS", "Show Filter", "Show GroupBox")
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
                Case "grdRSTSSPLX"
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

                    tlb_btn = DirectCast(tlb_pop.Tools("Export All to XLS"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ASCMAIN1.Running_in_VS)

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdRSTSSPLX.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Export All to XLS"
                If MsgBox("Export all SSPs to XLS?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                automated_XLS = True
                For Each grow As UltraWinGrid.UltraGridRow In grdRSTSSPLX.Rows
                    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                    Click_Command("View")
                    Click_Command("Save XLSX")
                    Click_Command("Done")
                    If CUST_CODE = "AMAZON" Then Exit For
                Next
                automated_XLS = False

                MsgBox("Export of all SSPs to XLS is Complete", MsgBoxStyle.OkOnly, "Success")

                Dim p As Process = Nothing
                p = Process.Start(ASCMAIN1.Folders("Work"))
                If p IsNot Nothing Then
                    p.Dispose()
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All for Channel"
                Dim CHANNEL_CODE As String = grd.ActiveRow.Cells("CHANNEL_CODE").Value & ""
                Dim sqlw As String = "CHANNEL_CODE = '" & CHANNEL_CODE & "'"
                For Each row As DataRow In dst.Tables("RSTSSPLX").Select(sqlw)
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
                If ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("HC_CODE").Text <> "" Then
                        If LookUp("ICTCOLL0", Absx1.txtFor("HC_CODE").Text) IsNot Nothing Then
                            XLS_Refresh_HC(Absx1.txtFor("HC_CODE").Text)
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
            Case "HC_CODE"
                If Absx1.txtFor("HC_CODE").Text <> "" Then
                    If LookUp("ICTCOLL0", Absx1.txtFor("HC_CODE").Text) IsNot Nothing Then
                        XLS_Refresh_HC(Absx1.txtFor("HC_CODE").Text)
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
        End Select
    End Sub

    Public Overrides Sub cbe_ValueChanged(sender As Object, e As EventArgs)
        MyBase.cbe_ValueChanged(sender, e)
        Refresh_Documents()
    End Sub

#End Region

    Sub Set_Month_Headings(SEASON_CODE As String)
        Dim SD As Date = CDate(IIf(SEASON_TYPE = "S", "02", "08") & "/01/" & SEASON_YEAR)

        Dim load_CUST_HC_HDGS As Boolean = (MONTH_HDGs.Count = 0)

        Dim rMOSHDG As Integer = XLSR("MOSHDG")
        For M As Integer = 0 To 6
            Dim D As Date = SD.AddMonths(M - 1)
            Dim LEGEND As String = Format(D, "MMM")
            workbook.Worksheets(0).Cells(rMOSHDG, c0 + M).Value = LEGEND
            If load_CUST_HC_HDGS Then MONTH_HDGs.Add(workbook.Worksheets(0).Cells(rMOSHDG, c0 + M).Value)
        Next

        Dim PRICE_BASE_DPCT As Decimal = 0

        Dim rCUST_CODE As Integer = XLSR("CUST_CODE")
        If blnConsolidated Then
            workbook.Worksheets(0).Cells(rCUST_CODE, 1).Value = "*"
            If CUST_CODES_consolidated.Count = 0 Then
                workbook.Worksheets(0).Cells(rCUST_CODE, 2).Value = "All Customers"
            Else
                workbook.Worksheets(0).Cells(rCUST_CODE, 2).Value = Join(CUST_CODES_consolidated.ToArray, ",")
            End If
        Else
            workbook.Worksheets(0).Cells(rCUST_CODE, 1).Value = CUST_CODE
            workbook.Worksheets(0).Cells(rCUST_CODE, 2).Value = Absx1.txtFor("CUST_NAME").Text
            PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            workbook.Worksheets(0).Cells(XLSR("TYPGRSNET"), 1).Value = 100 - PRICE_BASE_DPCT
        End If

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
        Fill_Records("RSTSSPLX", SEASON_CODE)
        Sort_grdColumns(grdRSTSSPLX, "CUST_CODE")
        grdRSTSSPLX.Text = "Stock & Sales Plans for " & SEASON_CODE

        Dim SF As String = "01"
        If Mid(SEASON_CODE, 5, 1) = "F" Then
            SF = "07"
        End If

        Dim TYP_BEG As String = Mid(SEASON_CODE, 1, 4) & SF
        Dim TYP_END As String = ASCMAIN1.Period_Calc(TYP_BEG, +6)
        lblTY.Text = $"TY Range: {TYP_BEG}-{TYP_END}"

        Dim LYP_BEG As String = ASCMAIN1.Period_Calc(TYP_BEG, -12)
        Dim LYP_END As String = ASCMAIN1.Period_Calc(TYP_END, -12)
        lblLY.Text = $"LY Range: {LYP_BEG}-{LYP_END}"

    End Sub

    Sub XLS_Validation(isTemplate As Boolean)
        Dim sheet_valid As Boolean = True
        Dim sheet_error_msg As String = ""
        XLSR.Clear()

        'If isTemplate Then Delete_Sheets.Clear()
        Delete_Sheets.Clear()

        xls_HC_CODEs.Clear()

        Dim LINE_TAGsAll As List(Of String) = {"PSIHDG", "TOAHDG", "MOSHDG", "ORGHDG", "LYAHDG", "CUST_CODE", "HC_CODE", "SEASON_CODE"}.ToList

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
                    xls_HC_CODEs.Add(sheet_name)
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

    Sub XLS_Refresh_HC(Optional single_HC_CODE As String = "")

        Dim rHC_CODE As Integer = XLSR("HC_CODE")

        If single_HC_CODE <> "" Then
            WorkbookView1.GetLock()
            ASCMAIN1.sql = "Select ICTCOLL0.* from ICTCOLL0 where HC_CODE = '" & single_HC_CODE & "'"
        Else
            ASCMAIN1.sql = "Select ICTCOLL0.* from ICTCOLL0,ICTBRAN1" & vbCrLf _
                    & " where ICTBRAN1.BRAND_CODE = ICTCOLL0.BRAND_CODE" & vbCrLf _
                    & "   and ICTCOLL0.HC_STATUS = 'A' and ICTBRAN1.BRAND_STATUS = 'A'"

            If Not blnConsolidated AndAlso rowSOTTCLS1.Item("AUTH_REQD") & "" = "1" Then
                ASCMAIN1.sql &= "" & vbCrLf _
                    & "   and ICTCOLL0.HC_CODE in (" & vbCrLf

                If Format(DTE2, "yyyyMMdd") > Format(Now.Date, "yyyyMMdd") Then ' If SEASON_is_current_or_future Then
                    ASCMAIN1.sql &= "" & vbCrLf _
                        & " (Select Distinct HC_CODE from SATAUTH1" & vbCrLf _
                        & "   where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & "     and OPS_YYYYPP_OPENED IS NOT NULL and OPS_YYYYPP_CLOSED IS NULL)" & vbCrLf _
                        & " UNION " & vbCrLf
                End If

                ASCMAIN1.sql &= "" & vbCrLf _
                    & " (Select Distinct HC_CODE from RSTSSPL2" & vbCrLf _
                    & " where SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                    & " and CUST_CODE = '" & CUST_CODE & "')" _
                    & " )"

            End If

            If blnConsolidated Then

                Dim sqlSEASON_CODE As String = " where SEASON_CODE = '" & SEASON_CODE & "'"
                If chk12Mos.Checked Then
                    sqlSEASON_CODE = $" where SEASON_CODE like '{Mid(SEASON_CODE, 1, 4)}%'"
                End If
                ASCMAIN1.sql &= "" & vbCrLf _
                   & "   and ICTCOLL0.HC_CODE in " & vbCrLf _
                   & " (Select Distinct HC_CODE from RSTSSPL2" & vbCrLf _
                   & sqlSEASON_CODE & vbCrLf _
                   & IIf(CUST_CODES_consolidated.Count = 0, "", " and CUST_CODE in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "') and (NVL(AMT_0,0) <> 0 or NVL(AMT_1,0) <> 0 or NVL(AMT_2,0) <> 0 or NVL(AMT_3,0) <> 0 or NVL(AMT_4,0) <> 0 or NVL(AMT_5,0) <> 0 or NVL(AMT_6,0) <> 0)" & vbCrLf) _
                   & " )"

            End If
        End If

        Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(workbook.Worksheets.Count - 1)

        If chk12Mos.Checked Then

            For XM As Integer = 1 To 6
                worksheet2.Cells(0, 11).EntireColumn.Insert(SpreadsheetGear.InsertShiftDirection.Right) ' HC template to be copied to each HC
                worksheet.Cells(0, 11).EntireColumn.Insert(SpreadsheetGear.InsertShiftDirection.Right) ' Totals
            Next

            Dim MOS() As String = {"", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan"}
            For XM As Integer = 1 To 6

                range = worksheet2.Cells(2, 10).EntireColumn
                range.Copy(worksheet2.Range(2, 10 + XM).EntireColumn, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                worksheet2.Cells(2, 10 + XM).Value = MOS(XM)

                range = worksheet.Cells(2, 10).EntireColumn
                range.Copy(worksheet.Range(2, 10 + XM).EntireColumn, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                worksheet.Cells(2, 10 + XM).Value = MOS(XM)

            Next



            Dim rc2 As Integer = worksheet.UsedRange.Rows.RowCount
            Dim xc As String = ""
            Dim F As String = ""
            For i As Integer = 0 To rc2
                xc = Excel_Cell0(i, 17) ' Col R
                F = worksheet.Cells(xc).Formula & ""
                If F.StartsWith("=SUM(") And F.EndsWith($"!L{CStr(i + 1)})") Then
                    F = F.Replace("!L", "!R")
                    worksheet.Cells(xc).Formula = F
                End If

                xc = Excel_Cell0(i, 18) ' Col S
                F = worksheet.Cells(xc).Formula & ""
                If F.StartsWith("=SUM(") And F.EndsWith($"!M{CStr(i + 1)})") Then
                    F = F.Replace("!M", "!S")
                    worksheet.Cells(xc).Formula = F
                End If
            Next

        End If

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "BRAND_CODE,HC_CODE")
            Dim HC_CODE As String = row.Item("HC_CODE")
            If Not xls_HC_CODEs.Contains(HC_CODE) Then
                Dim worksheet3 As SpreadsheetGear.IWorksheet = worksheet2
                For i As Integer = 1 To workbook.Worksheets.Count - 1
                    Dim sheet_name As String = workbook.Worksheets(i).Name
                    If Not Delete_Sheets.Contains(sheet_name) Then
                        If sheet_name > HC_CODE Then
                            worksheet3 = workbook.Worksheets(i)
                            Exit For
                        End If
                    End If
                Next
                worksheet = workbook.Worksheets.AddBefore(worksheet3)
                worksheet.Name = HC_CODE
                worksheet.Cells("E4").Activate()
                worksheet.WindowInfo.FreezePanes = True

                worksheet2.UsedRange.Copy(worksheet.Range("A1"), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

                For C As Integer = 0 To 50
                    worksheet.Cells(0, C).EntireColumn.ColumnWidth = worksheet2.Cells(0, C).EntireColumn.ColumnWidth
                Next

                Dim rowICTCOLL0 As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
                worksheet.Cells(rHC_CODE, 1).Value = HC_CODE
                worksheet.Cells(rHC_CODE, 2).Value = rowICTCOLL0.Item("HC_NAME")
                xls_HC_CODEs.Add(HC_CODE)
            End If
        Next

        If single_HC_CODE <> "" Then
            WorkbookView1.ReleaseLock()
            Absx1.txtFor("HC_CODE").Text = ""
        End If
    End Sub

    Sub XLS_Refresh_Allocations()

        ' TODO: GET ALLOCATIONS FOR ENTIRE YEAR

        'ASCMAIN1.sql = sqlSOTALLOX & vbCrLf _
        '    & "   and SOTALLO2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
        '    & "   and SOTALLO1.DATE_START between '" & Format(DTE1.AddMonths(-1), "dd-MMM-yyyy") & "' and '" & Format(DTE2.AddMonths(-1), "dd-MMM-yyyy") & "'" & vbCrLf _
        '    & "   and ICTCOLL1.HC_CODE in ('" & Join(xls_HC_CODEs.ToArray, "','") & "')"

        ' CHANGE BELOW IS SO THAT AN ALLOCATION THAT STRETCHES BETWEEN 06/01 THRU 07/31 WILL APPEAR ON BOTH THE SPRING AS WELL AS FALL SSPS
        ' DTE2X CHANGE BELOW IS FOR 12 MOS OPTION TO GET A FULL YEAR
        Dim DTE2X As Date = DTE2
        If chk12Mos.Checked Then DTE2X = DTE2X.AddMonths(6)
        Dim date_range As String = $" between '{Format(DTE1.AddMonths(-1), "dd-MMM-yyyy")}' and '{Format(DTE2X.AddMonths(-1), "dd-MMM-yyyy")}'"
        ASCMAIN1.sql = sqlSOTALLOX & vbCrLf _
            & $"   and SOTALLO2.CUST_CODE = '{CUST_CODE}'" & vbCrLf _
            & $"   and (SOTALLO1.DATE_START {date_range} or SOTALLO1.DATE_END {date_range})" & vbCrLf _
            & $"   and ICTCOLL1.HC_CODE in ('{Join(xls_HC_CODEs.ToArray, "','")}')"


        'If blnConsolidated Then
        '    ASCMAIN1.sql = Replace(Replace(Replace(Replace(ASCMAIN1.sql, "SOTALLO2.QTY_ALLO", "0 QTY_ALLO"), "SOTALLO1,SOTALLO2", "SOTALLO1"), "and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO", ""), "and SOTALLO2.CUST_CODE = ''", "")
        'Else
        ASCMAIN1.sql &= " UNION " & vbCrLf _
            & Replace(sqlSOTALLOX, "and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO", "and SOTALLO2.ALLO_CTL_NO (+) = SOTALLO1.ALLO_CTL_NO") & vbCrLf _
            & "   and SOTALLO2.CUST_CODE (+) = '" & CUST_CODE & "'" & vbCrLf _
            & "   and SOTALLO1.ALLO_CTL_NO in (Select Distinct LINE_KEY from RSTSSPL2" & vbCrLf _
            & " where RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "' and RSTSSPL2.CUST_CODE = '" & CUST_CODE & "' and LINE_TAG in ('TOAITM','PSIITM'))" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE in ('" & Join(xls_HC_CODEs.ToArray, "','") & "')"
        'End If

        If blnConsolidated Then
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(Replace(Replace(Replace(
                ASCMAIN1.sql, "SOTALLO2.QTY_ALLO", "0 QTY_ALLO"),
                "SOTALLO1,SOTALLO2", "SOTALLO1"),
                "and SOTALLO2.ALLO_CTL_NO = SOTALLO1.ALLO_CTL_NO", ""),
                "and SOTALLO2.CUST_CODE = ''", ""), "and SOTALLO2.CUST_CODE (+) = ''", ""),
                "and SOTALLO2.ALLO_CTL_NO (+) = SOTALLO1.ALLO_CTL_NO", ""),
            "and RSTSSPL2.CUST_CODE = '' ", "")
        End If


        Fill_Records("SOTALLOX", "", True, ASCMAIN1.sql)

        worksheet = workbook.Worksheets(0)

        Dim CA As Integer = 0
        If chk12Mos.Checked Then CA = 6

        Dim rTOAHDG As Integer = XLSR("TOAHDG")
        Dim rTOAITM As Integer = rTOAHDG + 2
        For Each row As DataRow In dst.Tables("SOTALLOX").Select("ITEM_SNU_CODE = 'N'", "ITEM_CODE,DATE_START")
            range = worksheet.Range(rTOAITM, 0).EntireRow
            rTOAITM += 1
            worksheet.Range(rTOAITM, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
            range.Copy(worksheet.Range(rTOAITM, 0), SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
            worksheet.Cells(rTOAITM, 1).Value = row.Item("ALLO_CTL_NO")   ' row.Item("ITEM_CODE")   'B
            worksheet.Cells(rTOAITM, 2).Value = row.Item("HC_CODE")     'C
            worksheet.Cells(rTOAITM, 3).Value = row.Item("ITEM_DESC")   'D
            worksheet.Cells(rTOAITM, 13 + CA).Value = row.Item("QTY_ALLO")   'N

            ' show entire shipping window, since we are now showing an allocation with a range of 06/01-07/31 in both Spring and Fall
            ' worksheet.Cells(rTOAITM, 14 + CAValue = row.Item("DATE_START") 'O
            worksheet.Cells(rTOAITM, 14 + CA).Value = Format(row.Item("DATE_START"), "MM/dd") & "-" & Format(row.Item("DATE_END"), "MM/dd") 'O

            worksheet.Cells(rTOAITM, 15 + CA).Value = row.Item("ITEM_CODE")  'P

            For Each HC_CODE As String In xls_HC_CODEs
                Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
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
            worksheet.Cells(rPSIITM, 2).Value = row.Item("HC_CODE")     'C
            worksheet.Cells(rPSIITM, 3).Value = row.Item("ITEM_DESC")   'D
            ' TODO: COL FOR ALL OF THE BELOW SHOULD MOVE 6 COLS TO THE RIGHT
            worksheet.Cells(rPSIITM, 13 + CA).Value = row.Item("QTY_ALLO")   'N
            worksheet.Cells(rPSIITM, 16 + CA).Value = row.Item("ITEM_RETAIL_PRICE")      'Q

            ' show entire shipping window, since we are now showing an allocation with a range of 06/01-07/31 in both Spring and Fall
            ' worksheet.Cells(rPSIITM, 17 + CA).Value = row.Item("DATE_START") 'R
            worksheet.Cells(rPSIITM, 17 + CA).Value = Format(row.Item("DATE_START"), "MM/dd") & "-" & Format(row.Item("DATE_END"), "MM/dd") 'R

            worksheet.Cells(rPSIITM, 18 + CA).Value = row.Item("ITEM_CODE")  'S
            worksheet.Cells(rPSIITM, 19 + CA).Value = row.Item("ITEM_CODE_COMPARE_TO")   'T
            worksheet.Cells(rPSIITM, 20 + CA).Value = row.Item("ITEM_CODE_COMPARE_TO_ALT")   'U


            If automated_Customer_HC_Pivot Then
                Dim rowRSTSSPP2 As DataRow = dst.Tables("RSTSSPP2").NewRow
                With rowRSTSSPP2
                    .Item("SEASON_CODE") = SEASON_CODE
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("ALLO_CTL_NO") = row.Item("ALLO_CTL_NO")
                    .Item("HC_CODE") = row.Item("HC_CODE")
                    .Item("ITEM_CODE") = row.Item("ITEM_CODE")
                    .Item("ITEM_DESC") = row.Item("ITEM_DESC")
                    .Item("ITEM_RETAIL_PRICE") = row.Item("ITEM_RETAIL_PRICE")
                    For M As Integer = 0 To 6
                        .Item("M0" & Format(M, "0")) = 0
                    Next
                End With
                dst.Tables("RSTSSPP2").Rows.Add(rowRSTSSPP2)
            End If

            For Each HC_CODE As String In xls_HC_CODEs
                Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
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
        For Each HC_CODE As String In xls_HC_CODEs
            Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
            For i As Integer = sLINE_TAG To eLINE_TAG
                worksheet2.Cells(i, 0).EntireRow.Hidden = (CStr(worksheet2.Cells(i, 2).Value & "") <> CStr(worksheet2.Cells(1, 1).Value & ""))
                If HC_CODE = xls_HC_CODEs(0) And i > sLINE_TAG Then XLS_Allocation_Lines.Add(worksheet2.Cells(i, 1).Value, i)
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
        For Each HC_CODE As String In xls_HC_CODEs
            Dim worksheet2 As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
            For i As Integer = sLINE_TAG To eLINE_TAG
                worksheet2.Cells(i, 0).EntireRow.Hidden = (CStr(worksheet2.Cells(i, 2).Value & "") <> CStr(worksheet2.Cells(1, 1).Value & ""))
                If HC_CODE = xls_HC_CODEs(0) And i > sLINE_TAG Then XLS_Allocation_Lines.Add(worksheet2.Cells(i, 1).Value, i)
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

    Sub Get_TY_Actuals(Optional ByVal blnFallPass As Boolean = False)

        Dim XLR As Int64 = 0
        Dim sql As String = ""

        ' TYA SLS (RETAIL)
        sql = ""
        For I As Integer = 0 To 6
            sql &= ", SUM(DECODE(RSTRETL1.OPS_YYYYPP,'" & YPs(I) & "',RSTRETL1.AMT_SOLD,0)) TYPSLSACT_" & CStr(I) & vbCrLf
        Next

        ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
            & sql _
            & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & IIf(blnConsolidated,
                  IIf(CUST_CODES_consolidated.Count = 0, "", " and RSTRETL1.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                  " and RSTRETL1.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf _
            & IIf(blnConsolidated, " and (RSTRETL1.CUST_CODE,ICTCOLL1.HC_CODE) in (Select CUST_CODE, HC_CODE from RSTSSPL2 where SEASON_CODE = '" & SEASON_CODE & "')", "") _
            & "   and RSTRETL1.OPS_YYYYPP between '" & YPs(0) & "' and '" & YPs(6) & "'" & vbCrLf _
            & "   group by ICTCOLL1.HC_CODE"
        Dim tblTYASLS As DataTable = ASCDATA1.GetDataTable
        tblTYASLS.PrimaryKey = New DataColumn() {tblTYASLS.Columns("HC_CODE")}


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

        ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
            & sql _
            & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & IIf(blnConsolidated,
                  IIf(CUST_CODES_consolidated.Count = 0, "", " and SOTINVH2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                  " and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf _
            & IIf(blnConsolidated, " and (SOTINVH2.CUST_CODE,ICTCOLL1.HC_CODE) in (Select CUST_CODE, HC_CODE from RSTSSPL2 where SEASON_CODE = '" & SEASON_CODE & "')", "") _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(0) & "' and '" & YPs(6) & "'" & vbCrLf _
            & "   group by ICTCOLL1.HC_CODE"
        Dim tblTYA As DataTable = ASCDATA1.GetDataTable
        tblTYA.PrimaryKey = New DataColumn() {tblTYA.Columns("HC_CODE")}


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

        ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
            & sql _
            & " from SOTORDR2,SOTORDR1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & IIf(blnConsolidated,
                  IIf(CUST_CODES_consolidated.Count = 0, "", " and SOTORDR2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                  " and SOTORDR2.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf _
            & IIf(blnConsolidated, " and (SOTORDR2.CUST_CODE,ICTCOLL1.HC_CODE) in (Select CUST_CODE, HC_CODE from RSTSSPL2 where SEASON_CODE = '" & SEASON_CODE & "')", "") _
            & "   and SOTORDR1.ORDR_STATUS >= 'O' and SOTORDR1.ORDR_STATUS <= 'P'" & vbCrLf _
            & "   group by ICTCOLL1.HC_CODE"

        Dim tblOP As DataTable = ASCDATA1.GetDataTable
        tblOP.PrimaryKey = New DataColumn() {tblOP.Columns("HC_CODE")}


        WorkbookView1.GetLock()

        'Debug.Print(workbook.Worksheets(2).Cells("R7").Formula)

        For Each HC_CODE As String In xls_HC_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
            If ws.ProtectContents Then
                ws.Unprotect(XLS_PWD)
            End If

            'Debug.Print(ws.Cells("R7").Formula & ":" & ws.Cells("S7").Formula)

            ' Clear out Actuals in case Open Order amount has moved, or history was corrected
            For I As Integer = 0 To 6
                ws.Cells(XLSR("TYAGRSB"), c0 + I).Value = 0
                ws.Cells(XLSR("TYAGRSP"), c0 + I).Value = 0
                ws.Cells(XLSR("TYAADJ"), c0 + I).Value = 0
            Next

            Dim rowTYASLS As DataRow = tblTYASLS.Rows.Find(HC_CODE)
            If rowTYASLS IsNot Nothing Then
                For I As Integer = 0 To 6
                    ws.Cells(XLSR("TYPSLSACT"), c0 + I).Value = Val(rowTYASLS.Item("TYPSLSACT_" & CStr(I)) & "") / 1000
                Next
            End If

            Dim rowTYA As DataRow = tblTYA.Rows.Find(HC_CODE)
            Dim rowOP As DataRow = tblOP.Rows.Find(HC_CODE)
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

                'If YPs(I) <= ASCMAIN1.CYP Then
                '    Dim f As String = ws.Cells(8, c0 + I).Formula
                '    '=IF(AND(E8>0,E30>0),E8/E30-1,"-")
                '    's/b
                '    '=IF(AND(E6>0,E30>0),E6/E30-1,"-")
                ' would also need to take care of totals sheet
                '    Dim CN As String = Excel_Cell0(XLSR("TYPSLS"), c0 + I) ' "E8"
                '    Dim CD As String = Excel_Cell0(XLSR("LYASLS"), c0 + I) ' "E30"
                '    f = String.Format("=IF(AND({0}>0,{1}>0),{0}/{1}-1,'-')", CN, CD)
                '    ws.Cells(8, c0 + I).Formula = Replace(f, "'", Chr(34))
                'End If
            Next
        Next

        WorkbookView1.ReleaseLock()
        If chk12Mos.Checked AndAlso Not blnFallPass AndAlso SEASON_CODE IsNot Nothing AndAlso SEASON_CODE.EndsWith("S") Then

            Dim saveSeason As String = SEASON_CODE
            Dim saveYPs(6) As String
            For i As Integer = 0 To 6
                saveYPs(i) = YPs(i)
            Next

            Dim saveC0 As Integer = c0

            Dim seasonFall As String = SEASON_CODE.Substring(0, 4) & "F"
            SEASON_CODE = seasonFall

            YPs(0) = seasonFall.Substring(0, 4) & "07"
            For i As Integer = 1 To 6
                YPs(i) = ASCMAIN1.Period_Calc(YPs(0), i)
            Next

            c0 = saveC0 + 6
            Get_TY_Actuals(True)

            c0 = saveC0
            SEASON_CODE = saveSeason
            For i As Integer = 0 To 6
                YPs(i) = saveYPs(i)
            Next
        End If

    End Sub
    Sub Get_LY_Actuals()

        ' This method should be called only once, when EntryMode = "N"

        Dim sql As String = ""

        Dim CUST_CODE_LY As String = CUST_CODE
        Dim rowRSTSSPLX As DataRow = LookUp("RSTSSPLX", New String() {SEASON_CODE, CUST_CODE})

        If rowRSTSSPLX IsNot Nothing Then
            CUST_CODE_LY = rowRSTSSPLX.Item("CUST_CODE_LY")
        End If

        ' LYA BOM

        sql = ""
        For I As Integer = 0 To 6
            sql &= ", SUM (RSTSSPL2.AMT_" & CStr(I) & ") LYABOM_" & CStr(I) & vbCrLf
        Next

        ASCMAIN1.sql = "Select RSTSSPL2.HC_CODE" & vbCrLf _
            & sql _
            & " from RSTSSPL2" & vbCrLf

        If EntryMode = "N" Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS
            ASCMAIN1.sql &= "" _
                & " where RSTSSPL2.LINE_TAG = 'TYPBOM'" & vbCrLf _
                & "   and RSTSSPL2.SEASON_CODE = '" & SEASON_CODE_LY & "'" & vbCrLf _
                & "   and RSTSSPL2.CUST_CODE = '" & CUST_CODE_LY & "'" & vbCrLf
        Else
            ASCMAIN1.sql &= "" _
                & " where RSTSSPL2.LINE_TAG = 'LYABOM'" & vbCrLf _
                & "   and RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & IIf(blnConsolidated,
                      IIf(CUST_CODES_consolidated.Count = 0, "", " and RSTSSPL2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                      " and RSTSSPL2.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf
        End If

        ASCMAIN1.sql &= "   group by RSTSSPL2.HC_CODE"

        Dim tblLYABOM As DataTable = ASCDATA1.GetDataTable
        tblLYABOM.PrimaryKey = New DataColumn() {tblLYABOM.Columns("HC_CODE")}

        ' LYA SLS (RETAIL)

        If EntryMode = "N" Or chkLY_Realtime.Checked Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS

            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM(DECODE(RSTRETL1.OPS_YYYYPP,'" & ASCMAIN1.Period_Calc(YPs(I), -12) & "',RSTRETL1.AMT_SOLD,0)) LYASLS_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
                & sql _
                & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & IIf(blnConsolidated,
                      IIf(CUST_CODES_consolidated.Count = 0, "", " and RSTRETL1.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                      " and RSTRETL1.CUST_CODE = '" & CUST_CODE_LY & "'") & vbCrLf _
                & IIf(blnConsolidated, " and (RSTRETL1.CUST_CODE,ICTCOLL1.HC_CODE) in (Select CUST_CODE, HC_CODE from RSTSSPL2 where SEASON_CODE = '" & SEASON_CODE & "')", "") _
                & "   and RSTRETL1.OPS_YYYYPP between '" & ASCMAIN1.Period_Calc(YPs(0), -12) & "' and '" & ASCMAIN1.Period_Calc(YPs(6), -12) & "'" & vbCrLf _
                & "   group by ICTCOLL1.HC_CODE"
        Else
            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM (1000 * RSTSSPL2.AMT_" & CStr(I) & ") LYASLS_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select RSTSSPL2.HC_CODE" & vbCrLf _
                & sql _
                & " from RSTSSPL2" & vbCrLf _
                & " where RSTSSPL2.LINE_TAG = 'LYASLS'" & vbCrLf _
                & "   and RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & IIf(blnConsolidated,
                        IIf(CUST_CODES_consolidated.Count = 0, "", " and RSTSSPL2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                        " and RSTSSPL2.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf _
                & "   group by RSTSSPL2.HC_CODE"
        End If

        Dim tblLYASLS As DataTable = ASCDATA1.GetDataTable
        tblLYASLS.PrimaryKey = New DataColumn() {tblLYASLS.Columns("HC_CODE")}

        ' LYA GRS & ADJ (@RETAIL)

        If EntryMode = "N" Or chkLY_Realtime.Checked Then ' OR IF ADDING A NEW HC - NEED TO CODE THIS
            Dim Z As String = "THEN SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_RETAIL_PRICE ELSE 0 END"
            sql = ""
            For I As Integer = 0 To 6
                Dim Y As String = "SOTINVH2.ORDR_YYYYPP_UPDATED='" & ASCMAIN1.Period_Calc(YPs(I), -12) & "'"
                sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'I' " & Z & ") LYAGRS_" & CStr(I) & vbCrLf
                sql &= ", SUM(CASE WHEN " & Y & " AND SOTINVH2.INV_TYPE = 'C' " & Z & ") LYAADJ_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select ICTCOLL1.HC_CODE" & vbCrLf _
                & sql _
                & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & IIf(blnConsolidated,
                      IIf(CUST_CODES_consolidated.Count = 0, "", " and SOTINVH2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                      " and SOTINVH2.CUST_CODE = '" & CUST_CODE_LY & "'") & vbCrLf _
                & IIf(blnConsolidated, " and (SOTINVH2.CUST_CODE,ICTCOLL1.HC_CODE) in (Select CUST_CODE, HC_CODE from RSTSSPL2 where SEASON_CODE = '" & SEASON_CODE & "')", "") _
                & "   and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & ASCMAIN1.Period_Calc(YPs(0), -12) & "' and '" & ASCMAIN1.Period_Calc(YPs(6), -12) & "'" & vbCrLf _
                & "   group by ICTCOLL1.HC_CODE"
        Else
            sql = ""
            For I As Integer = 0 To 6
                sql &= ", SUM (DECODE(RSTSSPL2.LINE_TAG,'LYAGRS',1000 * RSTSSPL2.AMT_" & CStr(I) & ",0)) LYAGRS_" & CStr(I) & vbCrLf
                sql &= ", SUM (DECODE(RSTSSPL2.LINE_TAG,'LYAADJ',1000 * RSTSSPL2.AMT_" & CStr(I) & ",0)) LYAADJ_" & CStr(I) & vbCrLf
            Next

            ASCMAIN1.sql = "Select RSTSSPL2.HC_CODE" & vbCrLf _
                & sql _
                & " from RSTSSPL2" & vbCrLf _
                & " where RSTSSPL2.LINE_TAG in ('LYAGRS','LYAADJ')" & vbCrLf _
                & "   and RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
                & IIf(blnConsolidated,
                        IIf(CUST_CODES_consolidated.Count = 0, "", " and RSTSSPL2.CUST_CODE  in ('" & Join(CUST_CODES_consolidated.ToArray, "','") & "')"),
                        " and RSTSSPL2.CUST_CODE = '" & CUST_CODE & "'") & vbCrLf _
                & "   group by RSTSSPL2.HC_CODE"
        End If

        Dim tblLYA As DataTable = ASCDATA1.GetDataTable
        tblLYA.PrimaryKey = New DataColumn() {tblLYA.Columns("HC_CODE")}

        WorkbookView1.GetLock()
        'Debug.Print(workbook.Worksheets(2).Cells("R7").Formula)
        For Each HC_CODE As String In xls_HC_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)

            'Debug.Print(ws.Cells("R7").Formula & ":" & ws.Cells("S7").Formula)

            Dim rowLYABOM As DataRow = tblLYABOM.Rows.Find(HC_CODE)
            If rowLYABOM IsNot Nothing Then
                For I As Integer = 0 To 0 '6
                    ws.Cells(XLSR("LYABOM"), 4 + I).Value = Val(rowLYABOM.Item("LYABOM_" & CStr(I)) & "")
                Next
            End If

            Dim rowLYASLS As DataRow = tblLYASLS.Rows.Find(HC_CODE)
            If rowLYASLS IsNot Nothing Then
                For I As Integer = 0 To 6
                    ws.Cells(XLSR("LYASLS"), 4 + I).Value = Val(rowLYASLS.Item("LYASLS_" & CStr(I)) & "") / 1000
                Next
            End If

            Dim rowLYA As DataRow = tblLYA.Rows.Find(HC_CODE)
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

        For Each HC_CODE As String In xls_HC_CODEs
            Dim ws As SpreadsheetGear.IWorksheet = workbook.Worksheets(HC_CODE)
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

    Private Sub grdRSTSSPLX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSSPLX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value & ""
            Click_Command("Edit")
        End If
    End Sub

    Sub Prepare_Summary_Budget()

        Dim O As Integer = IIf(SEASON_TYPE = "S", 0, 6)
        Dim sqlB As String = ""

        sqlB = ""
        For I As Integer = 1 To 6
            sqlB &= ", Sum (GLTACCT2.ACCT_BUD_P" & Format(I + O, "00") & ") P" & Format(I, "0") & vbCrLf
        Next
        Dim ACCT_CODE As String = "311000"
        ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
            & sqlB _
            & " from GLTACCT2,ICTCOLL1,SOTTCLS1" & vbCrLf _
            & " where GLTACCT2.ACCT_CODE = '" & ACCT_CODE & "'" & vbCrLf _
            & "   and GLTACCT2.ACCT_YEAR = '" & SEASON_YEAR & "'" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = DECODE(GLTACCT2.SEG3_CODE,'000','DPT',GLTACCT2.SEG3_CODE)" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = GLTACCT2.SEG4_CODE" & vbCrLf _
            & " group by SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CHANNEL_CODE As String = ROW.Item("CHANNEL_CODE")
            Dim HC_CODE As String = ROW.Item("HC_CODE")
            Dim rowSATBUDWX As DataRow = dst.Tables("SATBUDWX").NewRow
            rowSATBUDWX.Item("CHANNEL_CODE") = CHANNEL_CODE
            rowSATBUDWX.Item("HC_CODE") = HC_CODE
            For I As Integer = 1 To 6
                rowSATBUDWX("FBUD_" & Format(I, "0")) = -1 * Val(ROW.Item("P" & Format(I, "0")))
            Next
            dst.Tables("SATBUDWX").Rows.Add(rowSATBUDWX)
        Next

        ASCMAIN1.sql = "Select RSTSSPL2.*, ICTCOLL0.COLLECTION_CODE" & vbCrLf _
            & " from RSTSSPL2,ICTCOLL0" & vbCrLf _
            & " where RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
            & "   and RSTSSPL2.LINE_TAG in ('TYPGRSBNET','TYPGRSPNET')" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = RSTSSPL2.HC_CODE"
        RSTSSPL2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSSPL2 & " Add ITEM_BASIC_PROMO VARCHAR2(1)")
        ASCDATA1.ExecuteSQL("Update " & RSTSSPL2 & " Set ITEM_BASIC_PROMO = CASE WHEN LINE_TAG = 'TYPGRSBNET' THEN 'B' ELSE 'P' END")

        sqlB = ""
        For I As Integer = 1 To 6
            sqlB &= ", Sum (1000 * RSTSSPL2.AMT_" & Format(I - 1, "0") & ") P" & Format(I, "0") & vbCrLf
        Next
        ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, RSTSSPL2.HC_CODE" & vbCrLf _
            & sqlB _
            & " from " & RSTSSPL2 & " RSTSSPL2, ARTCUST1, SOTTCLS1" & vbCrLf _
            & "  where ARTCUST1.CUST_CODE = RSTSSPL2.CUST_CODE" & vbCrLf _
            & "    and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " group by SOTTCLS1.CHANNEL_CODE, RSTSSPL2.HC_CODE"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CHANNEL_CODE As String = ROW.Item("CHANNEL_CODE")
            Dim HC_CODE As String = ROW.Item("HC_CODE")
            Dim rowSATBUDWX As DataRow = dst.Tables("SATBUDWX").Rows.Find(New String() {CHANNEL_CODE, HC_CODE})
            If rowSATBUDWX Is Nothing Then
                rowSATBUDWX = dst.Tables("SATBUDWX").NewRow
                rowSATBUDWX.Item("CHANNEL_CODE") = CHANNEL_CODE
                rowSATBUDWX.Item("HC_CODE") = HC_CODE
                dst.Tables("SATBUDWX").Rows.Add(rowSATBUDWX)
            End If

            For I As Integer = 1 To 6
                rowSATBUDWX("SBUD_" & Format(I, "0")) = ROW.Item("P" & Format(I, "0"))
            Next

        Next







        ASCMAIN1.sql = "Select RSTSSPL2.*, ICTCOLL0.COLLECTION_CODE" & vbCrLf _
            & " from RSTSSPL2,ICTCOLL0" & vbCrLf _
            & " where RSTSSPL2.SEASON_CODE = '" & SEASON_CODE & "'" & vbCrLf _
            & "   and RSTSSPL2.LINE_TAG in ('TYPSLSACTZ')" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = RSTSSPL2.HC_CODE"
        Dim RSTSSPL2R As String = ASCMAIN1.Temp_Table

        sqlB = ""
        For I As Integer = 1 To 6
            sqlB &= ", Sum (1000 * RSTSSPL2R.AMT_" & Format(I - 1, "0") & ") P" & Format(I, "0") & vbCrLf
        Next
        ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, RSTSSPL2R.HC_CODE" & vbCrLf _
            & sqlB _
            & " from " & RSTSSPL2R & " RSTSSPL2R, ARTCUST1, SOTTCLS1" & vbCrLf _
            & "  where ARTCUST1.CUST_CODE = RSTSSPL2R.CUST_CODE" & vbCrLf _
            & "    and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " group by SOTTCLS1.CHANNEL_CODE, RSTSSPL2R.HC_CODE"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CHANNEL_CODE As String = ROW.Item("CHANNEL_CODE")
            Dim HC_CODE As String = ROW.Item("HC_CODE")
            Dim rowSATBUDWX As DataRow = dst.Tables("SATBUDWX").Rows.Find(New String() {CHANNEL_CODE, HC_CODE})
            If rowSATBUDWX Is Nothing Then
                rowSATBUDWX = dst.Tables("SATBUDWX").NewRow
                rowSATBUDWX.Item("CHANNEL_CODE") = CHANNEL_CODE
                rowSATBUDWX.Item("HC_CODE") = HC_CODE
                dst.Tables("SATBUDWX").Rows.Add(rowSATBUDWX)
            End If

            For I As Integer = 1 To 6
                rowSATBUDWX("RBUD_" & Format(I, "0")) = ROW.Item("P" & Format(I, "0"))
            Next

        Next




        Dim sqld As String = Get_SQL_SATBUDW1()

        ASCMAIN1.sql = "Select SATBUDW1.*,SOTTCLS1.CHANNEL_CODE,ICTCOLL1.HC_CODE" & vbCrLf _
            & " from SATBUDW1,ICTCOLL1,ARTCUST1,SOTTCLS1 where " & sqld & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SATBUDW1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SATBUDW1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE"
        Fill_Records("SATBUDW1", "", True, ASCMAIN1.sql)

        Dim SQLNOT As String = "" ' "CUST_CODE <> 'IPLBFIN'"
        ASCMAIN1.sql = "Select distinct CUST_CODE from SOTCHAN1 where CUST_CODE is Not Null"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CUST_CODE")
            SQLNOT &= " and CUST_CODE <> '" & row.Item("CUST_CODE") & "'"
        Next
        SQLNOT = Mid(SQLNOT, 6)

        For Each ROW As DataRow In dst.Tables("SATBUDW1").Select(SQLNOT)
            Dim CHANNEL_CODE As String = ROW.Item("CHANNEL_CODE")
            Dim HC_CODE As String = ROW.Item("HC_CODE")
            Dim rowSATBUDWX As DataRow = dst.Tables("SATBUDWX").Rows.Find(New String() {CHANNEL_CODE, HC_CODE})
            If rowSATBUDWX Is Nothing Then
                rowSATBUDWX = dst.Tables("SATBUDWX").NewRow
                rowSATBUDWX.Item("CHANNEL_CODE") = CHANNEL_CODE
                rowSATBUDWX.Item("HC_CODE") = HC_CODE
                dst.Tables("SATBUDWX").Rows.Add(rowSATBUDWX)
            End If

            For I As Integer = 1 To 6
                rowSATBUDWX("PBUD_" & Format(I, "0")) = Val(rowSATBUDWX("PBUD_" & Format(I, "0")) & "") + ROW.Item("WB_P" & Format(I + O, "00"))
            Next
        Next

        Sort_grdColumns(grdSATBUDWX, "HC_CODE")

        grdSATBUDWX.DisplayLayout.Bands(0).SortedColumns.Add("CHANNEL_CODE", False, True)

        grdSATBUDWX.Text = "Publish Budgets with Financial Adjustment for Season " & SEASON_CODE
    End Sub

    Overrides Function Excel_Import_Pre_Process_SG _
    (ByVal grd As UltraWinGrid.UltraGrid, dt As DataTable,
     Optional ByRef load_by_table As Boolean = False,
     Optional ByRef load_handled As Boolean = False,
     Optional ByRef F As ASFEXCL1 = Nothing) As Int64

        Dim dtbad As DataTable = dt.Clone
        dtbad.Columns.Add("ERROR")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading from XLS")

        Dim RowsMax As Int64 = dt.Rows.Count
        Dim r As Int64 = 0

        load_handled = True
        If dt.Rows.Count = 0 Then
            MsgBox("No Rows Loaded", MsgBoxStyle.OkOnly, "Import Failed")
        Else
            dst.Tables("RSTSSPL3").Rows.Clear()
        End If

        For Each row As DataRow In dt.Select("")
            r += 1
            If r Mod 100 = 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading data from XLS")
                RowsMax = dt.Rows.Count
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(RowsMax))
            End If

            Dim C As String = ""

            Try
                Dim rowRSTSSPL3 As DataRow = dst.Tables("RSTSSPL3").NewRow
                With rowRSTSSPL3
                    For Each C In New String() {"CUST_CODE", "ITEM_CODE"}
                        .Item(C) = row.Item(C)
                    Next

                    .Item("SEASON_CODE") = Absx1.cbeFor("SEASON_CODE").Value

                    For I As Integer = 1 To 6
                        C = "QTY_" & Format(I, "0")
                        .Item(C) = row.Item(C)
                        C = "AMT_" & Format(I, "0")
                        .Item(C) = row.Item(C)
                    Next
                End With

                dst.Tables("RSTSSPL3").Rows.Add(rowRSTSSPL3)

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

    Sub Box_Range()
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
    End Sub
    Sub Set_Data_Block(ByRef r As Integer, HC_CODE As String, tbl As DataTable, CHANNEL_CODE As String)

        r += 2

        Dim ST As Integer = 0
        If chk12Mos.Checked Then ST = 6


        If HC_CODE = "" Then
            worksheet.Cells(r, c0 - 1).Value = "Total"
        Else
            worksheet.Cells(r, c0 - 1).Value = HC_CODE
        End If
        For i As Integer = 0 To 6 + ST + 2
            worksheet.Cells(r, c0 + i).Formula = "=Total!" & Excel_Cell0(2, c0 + i)
        Next
        range = worksheet.Cells(r, c0 - 1, r, c0 + 6 + ST + 2)
        range.Font.Bold = True
        Box_Range()

        If chk12Mos.Checked Then
            Dim c00 As Integer = c0 + 6 + ST + 2 + 1
            Dim c00_orig As Integer = c00
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q1"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q2"
            c00 += 1 : worksheet.Cells(r, c00).Value = "S 445 Total"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q1"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q2"
            c00 += 1 : worksheet.Cells(r, c00).Value = "S 454 Total"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q3"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q4"
            c00 += 1 : worksheet.Cells(r, c00).Value = "F 445 Total"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q3"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Q4"
            c00 += 1 : worksheet.Cells(r, c00).Value = "F 454 Total"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Total Year 445"
            c00 += 1 : worksheet.Cells(r, c00).Value = "Total Year 454"

            range = worksheet.Cells(r, c00_orig + 1, r, c00)
            range.Font.Bold = True
            Box_Range()
            With range.EntireColumn
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                .ColumnWidth = 12
                .NumberFormat = "#,##0.0;(#,##0.0)"
            End With

            range = worksheet.Cells(r, c00 - 1, r, c00)
            With range.EntireColumn
                .ColumnWidth = 15
            End With
        End If



        If SEASON_TYPE = "F" Then

            worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = SEASON_CODE_prior & " 445"
            worksheet.Cells(r, c0 + 6 + 2 + 1 + 2).Value = "Total Year"

            range = worksheet.Cells(r, c0 + 6 + 2 + 1 + 1, r, c0 + 6 + 2 + 1 + 2)
            range.Font.Bold = True
            Box_Range()
        End If

        Dim r0 As Integer = r

        r += 1 : Set_Data_Block_row(r, 7, HC_CODE) 'TYPSLSACT
        r += 1 : Set_Data_Block_row(r, 8, HC_CODE) 'TYPSLSACTZ

        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Financial Retail Plan" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Blue
        'r += 1 : worksheet.Cells(r, c0 - 1).Value = "Carryover" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Red

        r += 1 : Set_Data_Block_row(r, 10, HC_CODE) ' TYPGRSB
        r += 1 : Set_Data_Block_row(r, 11, HC_CODE) ' TYPGRSP
        r += 1 : Set_Data_Block_row(r, 12, HC_CODE) ' TYPGRS
        r += 1 : Set_Data_Block_row(r, 14, HC_CODE) ' TYPADJ
        r += 1 : Set_Data_Block_row(r, 15, HC_CODE) ' TYPGRSBNET
        r += 1 : Set_Data_Block_row(r, 16, HC_CODE) ' TYPGRSPNET
        r += 1 : Set_Data_Block_row(r, 17, HC_CODE) : worksheet.Cells(r, c0 - 1, r, c0 + 6 + ST + 2).Interior.Color = SpreadsheetGear.Colors.LightGray ' TYPGRSNET
        If SEASON_TYPE = "F" Then
            worksheet.Cells(r, c0 + 6 + 2 + 1 + 1, r, c0 + 6 + 2 + 1 + 2).Interior.Color = SpreadsheetGear.Colors.LightGray
        End If
        If chk12Mos.Checked Then
            Dim c00 As Integer = c0 + 6 + ST + 2 + 1
            worksheet.Cells(r, c00 + 1, r, c00 + 12).Interior.Color = SpreadsheetGear.Colors.LightGray
        End If

        If CHANNEL_CODE = "" And Not chk12Mos.Checked Then
        Else
            r += 1 : worksheet.Cells(r, c0 - 1).Value = "Financial Gross Plan" : worksheet.Cells(r, c0 - 1).EntireRow.Font.Color = SpreadsheetGear.Colors.Blue

            Dim BUD_PS As Decimal = 0
            Dim FBUD() As Decimal
            ReDim FBUD(6 + ST)
            If HC_CODE = "Total" Then
                For i As Integer = 0 To 6 + ST
                    FBUD(i) = Val(tbl.Compute("SUM(P" & Format(i, "0") & ")", "CHANNEL_CODE='1'") & "") / 1000
                Next
            Else
                Dim row As DataRow = tbl.Rows.Find(New String() {"1", HC_CODE})
                If row IsNot Nothing Then
                    For i As Integer = 0 To 6 + ST
                        FBUD(i) = Val(row.Item("P" & Format(i, "0")) & "") / 1000
                    Next
                    If SEASON_TYPE = "F" Then
                        BUD_PS = Val(row.Item("BUD_PS") & "")
                    End If
                End If
            End If
            For i As Integer = 0 To 6 + ST
                worksheet.Cells(r, c0 + i).Value = FBUD(i)
            Next

            worksheet.Cells(r, c0 + 6 + ST + 1).Formula = "=SUM(" & Excel_Cell0(r, c0 + 0) & ":" & Excel_Cell0(r, c0 + 5 + ST) & ")"
            worksheet.Cells(r, c0 + 6 + ST + 2).Formula = "=SUM(" & Excel_Cell0(r, c0 + 1) & ":" & Excel_Cell0(r, c0 + 6 + ST) & ")"

            If SEASON_TYPE = "F" And HC_CODE <> "" Then
                'Dim row As DataRow = dst.Tables("RSTSSPLR").Rows.Find(New String() {HC_CODE, LINE_TAG})
                'If row IsNot Nothing Then
                '    worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = Val(row.Item("AMT") & "")
                'End If
                worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = BUD_PS / 1000
            End If

            Set_Data_Block_qtr(r, ST)
        End If

        r += 1 : Set_Data_Block_row(r, 18, HC_CODE) ' TYPEOM
        worksheet.Cells(r, c0 + 6 + ST + 1).Value = DBNull.Value  ' EOM
        worksheet.Cells(r, c0 + 6 + ST + 2).Value = DBNull.Value  ' EOM


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

    Sub Set_Data_Block_row(r As Integer, rSource As Integer, HC_CODE As String)

        Dim LINE_TAG As String = workbook.Worksheets("Total").Cells(rSource, 0).Value
        Dim SHEET_NAME As String = HC_CODE
        If SHEET_NAME = "" Then SHEET_NAME = "Total"

        Dim ST As Integer = 0
        If chk12Mos.Checked Then ST = 6

        worksheet.Cells(r, c0 - 1).Formula = "=" & SHEET_NAME & "!" & Excel_Cell0(rSource, c0 - 1)
        For i As Integer = 0 To 6 + ST + 2 ' TODO: WE MAY NEED TO EMBELLISH THIS A BIT WITH THE QUARTERS; BUT FIRST LET'S FIND OUT WHAT WE ARE DOING TO THE SHEETS
            worksheet.Cells(r, c0 + i).Formula = "='" & SHEET_NAME & "'!" & Excel_Cell0(rSource, c0 + i)
        Next

        If worksheet.Name.Equals("Roll-Up", StringComparison.OrdinalIgnoreCase) Then
            worksheet.Cells(r, 0).Value = RollUpCustLabel
        End If

        If HC_CODE <> "" AndAlso Not HC_CODE.Equals("Total", StringComparison.OrdinalIgnoreCase) Then
            worksheet.Cells(r, 1).Value = HC_CODE
        Else
            worksheet.Cells(r, 1).Value = Nothing   ' no "Total" in col B
        End If
        If SEASON_TYPE = "F" And HC_CODE <> "" Then
            Dim row As DataRow = dst.Tables("RSTSSPLR").Rows.Find(New String() {HC_CODE, LINE_TAG})
            If row IsNot Nothing Then
                worksheet.Cells(r, c0 + 6 + 2 + 1 + 1).Value = Val(row.Item("AMT") & "")
            End If
        End If

        Set_Data_Block_qtr(r, ST)



    End Sub

    Sub Set_Data_Block_qtr(r As Int32, st As Int32)
        If chk12Mos.Checked Then
            Dim c00 As Integer = c0 + 6 + st + 2 + 1
            Dim q0 As Integer = 3
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 1)}:{Excel_Cell0(r, q0 + 3)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 4)}:{Excel_Cell0(r, q0 + 6)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, c00 - 2)}:{Excel_Cell0(r, c00 - 1)})"
            Dim q1 As Integer = c00
            'q0 += 3
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 2)}:{Excel_Cell0(r, q0 + 4)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 5)}:{Excel_Cell0(r, q0 + 7)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, c00 - 2)}:{Excel_Cell0(r, c00 - 1)})"
            Dim q2 As Integer = c00
            q0 += 3 + 3
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 1)}:{Excel_Cell0(r, q0 + 3)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 4)}:{Excel_Cell0(r, q0 + 6)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, c00 - 2)}:{Excel_Cell0(r, c00 - 1)})"
            Dim q3 As Integer = c00
            'q0 += 3
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 2)}:{Excel_Cell0(r, q0 + 4)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, q0 + 5)}:{Excel_Cell0(r, q0 + 7)})"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"=SUM({Excel_Cell0(r, c00 - 2)}:{Excel_Cell0(r, c00 - 1)})"
            Dim q4 As Integer = c00

            q0 += 3
            'c00 += 1 : worksheet.Cells(r, c00).Formula = $"={Excel_Cell0(r, q1)}+{Excel_Cell0(r, q2)}+{Excel_Cell0(r, q3)}+{Excel_Cell0(r, q4)}"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"={Excel_Cell0(r, q1)}+{Excel_Cell0(r, q3)}"
            c00 += 1 : worksheet.Cells(r, c00).Formula = $"={Excel_Cell0(r, q2)}+{Excel_Cell0(r, q4)}"
        End If
    End Sub

    Sub Prepare_RollUp()

        ' If Not ASCMAIN1.Running_in_VS Then Exit Sub

        Dim CHANNEL_CODE As String = "1"
        For Each row As DataRow In dst.Tables("RSTSSPLX").Select("")
            If row.Item("CHANNEL_CODE") & "" = "1" And row.Item("SEL") & "" <> "1" Then ' a Channel 1 customer was NOT selected
                CHANNEL_CODE = ""
                ' Exit Sub
            End If
            If row.Item("CHANNEL_CODE") & "" <> "1" And row.Item("SEL") & "" = "1" Then ' a non-Channel 1 customer WAS selected
                CHANNEL_CODE = ""
                ' Exit Sub
            End If
        Next

        Dim custList As New List(Of String)
        For Each rowX As DataRow In dst.Tables("RSTSSPLX").Select("SEL = '1'", "CUST_CODE")
            Dim c As String = rowX.Item("CUST_CODE") & ""
            If c <> "" AndAlso Not custList.Contains(c) Then
                custList.Add(c)
            End If
        Next
        RollUpCustLabel = String.Join(", ", custList.ToArray())

        Dim O As Integer = IIf(SEASON_TYPE = "S", 0, 6)
        Dim sqlB As String = ""

        Dim SEASON_YEAR_NY As String = Format(Val(SEASON_YEAR) + 1, "0000")

        sqlB = ""
        For I As Integer = 1 To If(chk12Mos.Checked, 13, 7)
            If (SEASON_TYPE = "F" And I = 7) Or (I = 13) Then
                sqlB &= ", Sum (-1 * Decode(GLTACCT2.ACCT_YEAR,'" & SEASON_YEAR_NY & "',GLTACCT2.ACCT_BUD_P01,0)) P" & Format(I - 1, "0") & vbCrLf
            Else
                sqlB &= ", Sum (-1 * Decode(GLTACCT2.ACCT_YEAR,'" & SEASON_YEAR & "',GLTACCT2.ACCT_BUD_P" & Format(I + O, "00") & ",0)) P" & Format(I - 1, "0") & vbCrLf
            End If
        Next

        If SEASON_TYPE = "F" Then
            Dim B As String = ""
            For I As Integer = 1 To 6
                B &= "+NVL(GLTACCT2.ACCT_BUD_P" & Format(I, "00") & ",0)"
            Next
            sqlB &= ", Sum (-1 * Decode(GLTACCT2.ACCT_YEAR,'" & SEASON_YEAR & "'," & Mid(B, 2) & ",0)) BUD_PS" & vbCrLf
        End If
        Dim ACCT_CODE As String = "311000"
        ASCMAIN1.sql = "Select SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
            & sqlB _
            & " from GLTACCT2,ICTCOLL1,SOTTCLS1" & vbCrLf _
            & " where GLTACCT2.ACCT_CODE = '" & ACCT_CODE & "'" & vbCrLf _
            & "   and (GLTACCT2.ACCT_YEAR = '" & SEASON_YEAR & "' or GLTACCT2.ACCT_YEAR = '" & SEASON_YEAR_NY & "')" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = DECODE(GLTACCT2.SEG3_CODE,'000','DPT',GLTACCT2.SEG3_CODE)" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = GLTACCT2.SEG4_CODE" & vbCrLf _
            & " group by SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        tbl.PrimaryKey = New DataColumn() {tbl.Columns("CHANNEL_CODE"), tbl.Columns("HC_CODE")}

        WorkbookView1.GetLock()
        worksheet = workbook.Worksheets.AddAfter(workbook.Worksheets("Total"))
        worksheet.Name = "Roll-Up"

        Dim rrow As Integer = 0

        Dim mPLUS As Integer = 0

        'Stop ' TODO: ADD 6 COLS
        Dim m12 As Integer = 0
        If chk12Mos.Checked Then m12 = 6



        If SEASON_TYPE = "F" Then mPLUS = 1 + 2

        For m As Integer = -2 To 6 + m12 + 2 + mPLUS
            If m = 6 + m12 + 2 + 1 Then
                ' DO NOTHING - THIS IS THE SPACE BETWEEN THE F SEASON AND S
            Else
                With worksheet.Cells(0, c0 + m).EntireColumn
                    .ColumnWidth = workbook.Worksheets("Total").Cells(0, c0 + m).EntireColumn.ColumnWidth
                    If m >= 0 Then
                        .NumberFormat = "#,##0.0;(#,##0.0)"
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    End If
                End With
            End If
        Next

        Dim ALL_HCs As String = ""
        Set_Data_Block(rrow, "Total", tbl, CHANNEL_CODE)
        For Each HC_CODE In xls_HC_CODEs
            ALL_HCs &= "+" & Excel_Cell0(rrow + 2 + 1, c0 + 6 + m12 + 2 + 1 + 1)
            Set_Data_Block(rrow, HC_CODE, tbl, CHANNEL_CODE)
        Next

        If SEASON_TYPE = "F" Then
            Dim m0 As Integer = c0 + 6 + 2 + 1
            Dim r0 As Integer = 3
            worksheet.Cells(r0 + 0, m0 + 1).Formula = "=" & Mid(ALL_HCs, 2)

            Dim rx As Integer = 6
            If CHANNEL_CODE = "" Then rx = 5

            range = worksheet.Cells(Excel_Cell0(r0 + 0, m0 + 1) & ":" & Excel_Cell0(r0 + rx, m0 + 1))
            worksheet.Cells(r0 + 0, m0 + 1).Copy(range, SpreadsheetGear.PasteType.FormulasAndNumberFormats, SpreadsheetGear.PasteOperation.None, False, False)
            worksheet.Cells(r0 + rx, m0 + 1).Value = DBNull.Value
            worksheet.Cells(r0 + rx, m0 + 2).Value = DBNull.Value
        End If

        WorkbookView1.ReleaseLock()


    End Sub

    Function Excel_Import_SG_from_Forecast_File() As Int64

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Workbook to Import"
            ' openFileDialog1.Filter = "xls files (*.xls)|*.xls" ' |xlsx files (*.xlsx)|*.xlsx"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.RestoreDirectory = True

            Excel_Import_SG_from_Forecast_File = -1

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If FILENAME <> "" Then
            Try
                oWB = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                If oWB.Worksheets.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Function
                End If

                Dim EXCEL_SHEET As String = ""
                For ws As Integer = 0 To oWB.Worksheets.Count - 1
                    If oWB.Worksheets(ws).Name = "Data" Then
                        EXCEL_SHEET = oWB.Worksheets(ws).Name
                        Exit For
                    End If
                Next
                If EXCEL_SHEET = "" Then
                    MsgBox("No Data Sheet Found")
                    Exit Function
                End If
                oSheet = oWB.Sheets(EXCEL_SHEET)

                dst.Tables("RSTSSPL3").Rows.Clear()

                Dim rCount As Int64 = oSheet.UsedRange.RowCount
                Dim rows_added As Int64 = -1

                Dim SEASON_CODE As String = Absx1.cbeFor("SEASON_CODE").Value

                Dim PX As Integer = 15
                Dim SX As Integer = 0
                If SEASON_CODE.EndsWith("F") Then SX = 30

                For r As Int64 = 2 To rCount - 1
                    Dim CUST_CODE As String = oSheet.Cells(r, 2).Text
                    Dim ITEM_CODE As String = oSheet.Cells(r, 8).Text

                    Dim has_data As Boolean = False
                    If CUST_CODE <> "" And ITEM_CODE <> "" Then
                        Dim rowRSTSSPL3 As DataRow = dst.Tables("RSTSSPL3").NewRow
                        rowRSTSSPL3.Item("CUST_CODE") = CUST_CODE
                        rowRSTSSPL3.Item("ITEM_CODE") = ITEM_CODE
                        rowRSTSSPL3.Item("SEASON_CODE") = SEASON_CODE
                        For c As Integer = 1 To 6
                            Dim RX As Integer = PX + SX + (c - 1) * 5
                            Dim QTY As Int64 = Val(oSheet.Cells(r, RX + 0).Value & "")
                            Dim AMT As Decimal = Val(oSheet.Cells(r, RX + 1).Value & "") ' * 1000
                            'If QTY <> 0 Or AMT <> 0 Then - NEEDED TO RECORD RECORDS WITH 0 OR ELSE KOHLS 2018F BUDGETS WOULD NOT GET LOADED
                            rowRSTSSPL3.Item("QTY_" & Format(c, "0")) = QTY
                            rowRSTSSPL3.Item("AMT_" & Format(c, "0")) = AMT
                            has_data = True
                            'End If
                        Next

                        If has_data Then
                            dst.Tables("RSTSSPL3").Rows.Add(rowRSTSSPL3)
                            If rows_added = -1 Then rows_added = 0
                            rows_added += 1
                        End If
                    End If
                Next

                Excel_Import_SG_from_Forecast_File = rows_added

            Catch ex As Exception
                MsgBox("Exception Occurred:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error Opening Excel Workbook")
            Finally

            End Try

        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Application.DoEvents()
    End Function

    Sub Customer_HC_Pivot()

        dst.Tables("RSTSSPP1").Rows.Clear()
        dst.Tables("RSTSSPP2").Rows.Clear()
        MONTH_HDGs = New List(Of String)

        automated_Customer_HC_Pivot = True

        automated_XLS = True

        For Each CUST_CODE As String In CUST_CODES_consolidated
            If CUST_CODE = "AMAZON" Then
                ' no report for AMAZON
            Else
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("View")
                Click_Command("Save XLSX")

                Click_Command("Done")
            End If
        Next

        Dim range As SpreadsheetGear.IRange = Nothing
        Dim XLS_FILENAME As String = ""

        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing

        oSheet = oWB.Worksheets(0)
        oSheet.Name = "Customer HC"
        oSheet.Range(0, 0).CopyFromDataTable(dst.Tables("RSTSSPP1"), SpreadsheetGear.Data.SetDataFlags.None)
        For m As Integer = 0 To 6
            Dim c As Integer = 5 + m

            With oSheet.Range(0, c).EntireColumn
                .NumberFormat = "#,##0.0"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
            oSheet.Cells(0, c).Value = "$" & MONTH_HDGs(m)
            oSheet.Cells(0, c).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        Next
        With oSheet.Range(0, 3).EntireColumn
            .NumberFormat = "#,##0"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With
        oSheet.Cells("A:Z").Columns.AutoFit()
        For C As Integer = 0 To 25
            If C >= 5 Then
                oSheet.Cells(0, C).EntireColumn.ColumnWidth = 8
            Else
                oSheet.Cells(0, C).EntireColumn.ColumnWidth = oSheet.Cells(0, C).EntireColumn.ColumnWidth * 1.15
            End If
        Next

        oSheet = oWB.Worksheets.Add()
        oSheet.Name = "Promo Saleable"
        With oSheet.Range(0, 2).EntireColumn
            .NumberFormat = "@"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With
        oSheet.Range(0, 0).CopyFromDataTable(dst.Tables("RSTSSPP2"), SpreadsheetGear.Data.SetDataFlags.None)
        For m As Integer = 0 To 6
            Dim c As Integer = 7 + m

            With oSheet.Range(0, c).EntireColumn
                .NumberFormat = "#,##0.0"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
            oSheet.Cells(0, c).Value = "$" & MONTH_HDGs(m)
            oSheet.Cells(0, c).HorizontalAlignment = SpreadsheetGear.HAlign.Left

            c += 7 ' units columns
            With oSheet.Range(0, c).EntireColumn
                .NumberFormat = "#,##0.0"
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With
            oSheet.Cells(0, c).Value = "#" & MONTH_HDGs(m)
            oSheet.Cells(0, c).HorizontalAlignment = SpreadsheetGear.HAlign.Left

            ' units column values
            Dim ru As Integer = 0
            Do While oSheet.Cells(ru + 1, 0).Value & "" <> ""
                ru += 1
                Dim PRICE As Decimal = Val(oSheet.Cells(ru, 6).Value & "")
                Dim AMT As Decimal = Val(oSheet.Cells(ru, c - 7).Value & "")

                If AMT <> 0 And PRICE <> 0 Then
                    oSheet.Cells(ru, c).Value = 1000 * AMT / PRICE
                End If
            Loop
        Next


        With oSheet.Range(0, 2).EntireColumn
            .NumberFormat = "@"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With oSheet.Range(0, 6).EntireColumn
            .NumberFormat = "#,##0.00"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With

        oSheet.Cells("A:Z").Columns.AutoFit()
        For C As Integer = 0 To 25
            If C >= 7 Then
                oSheet.Cells(0, C).EntireColumn.ColumnWidth = 8
            Else
                oSheet.Cells(0, C).EntireColumn.ColumnWidth = oSheet.Cells(0, C).EntireColumn.ColumnWidth * 1.15
            End If
        Next


        Dim CHCP As String = ASCMAIN1.Next_Control_No(Me.Name & ".CUSTOMER_HC_PIVOT")
        XLS_FILENAME = ASCMAIN1.Folders("Work") & "Customer_HC_Pivot_" & CHCP & ".xlsx"
        oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)



        automated_XLS = False
        automated_Customer_HC_Pivot = False

        dst.Tables("RSTSSPP1").Rows.Clear()
        dst.Tables("RSTSSPP2").Rows.Clear()

        MsgBox("Customer HC Pivot XLS is Complete", MsgBoxStyle.OkOnly, "Success")

        Dim p As Process = Nothing
        p = Process.Start(ASCMAIN1.Folders("Work"))
        If p IsNot Nothing Then
            p.Dispose()
        End If


    End Sub

    Private Sub grdRSTSSPLX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTSSPLX.InitializeLayout

    End Sub

    Private Sub grdRSTSSPLX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdRSTSSPLX.InitializeRow
        If e.Row.IsDataRow Then
            Dim UPD As String = e.Row.Cells("UPD").Value & ""
            If UPD = "1" Then
                e.Row.Cells("UPD").Appearance.BackColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("UPD").Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End If


    End Sub
    Private Sub Get_Retailer_BOM_EOW()
        If Not XLSR.ContainsKey("RTLOHB") Then Exit Sub
        WorkbookView1.GetLock()
        Try
            For m As Integer = 0 To 6
                Dim yp As String = YPs(m)
                Dim prevYP As String = ASCMAIN1.Period_Calc(yp, -1)

                ASCMAIN1.sql =
                "select ICTCOLL1.HC_CODE, " & vbCrLf &
                "       SUM(R.QTY_EOW * I.ITEM_RETAIL_PRICE / 1000) AMT" & vbCrLf &
                "  from RSTRETL1 R, ICTITEM1 I, ICTCOLL1, GLTPARM3 G" & vbCrLf &
                " where I.ITEM_CODE = R.ITEM_CODE" & vbCrLf &
                "   and ICTCOLL1.COLLECTION_CODE = I.COLLECTION_CODE" & vbCrLf &
                IIf(blnConsolidated,
                   IIf(CUST_CODES_consolidated.Count = 0, "",
                      "   and R.CUST_CODE in ('" & String.Join("','", CUST_CODES_consolidated.ToArray) & "')" & vbCrLf),
                   "   and R.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf) &
                "   and G.YYYYPP = '" & prevYP & "'" & vbCrLf &
                "   and G.REL_WEEK = G.MAX_WEEK" & vbCrLf &
                "   and R.OPS_YYYYWW = G.YYYYWW" & vbCrLf &
                " group by ICTCOLL1.HC_CODE"

                Dim tbl As DataTable = ASCDATA1.GetDataTable()
                Dim hasAmt As Boolean = (tbl.Columns.Contains("AMT"))
                tbl.PrimaryKey = New DataColumn() {tbl.Columns("HC_CODE")}

                Dim rRTLOHB As Integer = XLSR("RTLOHB")
                Dim rSales As Integer = XLSR("TYPSLSACTZ")
                Dim rShip As Integer = XLSR("TYPGRS")
                Dim rAdj As Integer = XLSR("TYPADJ")

                Dim wsTotal = workbook.Worksheets(0)

                For Each HC_CODE As String In xls_HC_CODEs
                    Dim ws = workbook.Worksheets(HC_CODE)
                    Dim row As DataRow = Nothing
                    If tbl IsNot Nothing AndAlso tbl.Rows IsNot Nothing AndAlso tbl.Rows.Count > 0 Then
                        row = tbl.Rows.Find(HC_CODE)
                    End If

                    TryUnprotect(ws)
                    Try
                        Dim target = ws.Cells(rRTLOHB, c0 + m)

                        Dim hasSqlAmt As Boolean = False
                        Dim sqlAmt As Decimal = 0D
                        If row IsNot Nothing AndAlso hasAmt Then
                            Dim raw = row("AMT")
                            If raw IsNot Nothing AndAlso Not IsDBNull(raw) AndAlso raw.ToString() <> "" Then
                                sqlAmt = Convert.ToDecimal(raw)
                                hasSqlAmt = True
                            End If
                        End If

                        If hasSqlAmt Then
                            target.ClearContents()
                            target.Value = sqlAmt
                        Else
                            If m > 0 Then
                                Dim p As Integer = c0 + (m - 1)
                                Dim addrPrev = ws.Cells(rRTLOHB, p).Address.Replace("$", "")         'E6
                                Dim addrPrevAbove = ws.Cells(rRTLOHB - 1, p).Address.Replace("$", "") 'E5
                                Dim addrS = ws.Cells(rSales, p).Address.Replace("$", "")             'E9
                                Dim addrG = ws.Cells(rShip, p).Address.Replace("$", "")              'E13
                                Dim addrA = ws.Cells(rAdj, p).Address.Replace("$", "")               'E15

                                If m = 1 Then
                                    'First “data” month (F column): IF(E6="", E5-E9+E13+E15, E6-E9+E13+E15)
                                    target.Formula = "=IF(" & addrPrev & "="""", " &
                              addrPrevAbove & "-" & addrS & "+" & addrG & "+" & addrA &
                              ", " & addrPrev & "-" & addrS & "+" & addrG & "+" & addrA & ")"
                                Else
                                    'Later months: standard rolling formula
                                    target.Formula = "=" & addrPrev & "-" & addrS & "+" & addrG & "+" & addrA
                                End If
                            Else
                                target.Clear()
                            End If
                        End If
                    Finally
                        TryProtect(ws)
                    End Try
                Next

                TryUnprotect(wsTotal)
                Try
                    Dim parts As New List(Of String)
                    For Each HC_CODE As String In xls_HC_CODEs
                        Dim hcAddr = workbook.Worksheets(HC_CODE).Cells(rRTLOHB, c0 + m).Address.Replace("$", "")
                        parts.Add("'" & HC_CODE & "'!" & hcAddr)
                    Next
                    wsTotal.Cells(rRTLOHB, c0 + m).Formula = "=SUM(" & String.Join(",", parts) & ")"
                Finally
                    TryProtect(wsTotal)
                End Try
            Next

        Finally
            WorkbookView1.ReleaseLock()
        End Try
    End Sub
    Private Sub TryUnprotect(ws As SpreadsheetGear.IWorksheet)
        Try : ws.Unprotect(XLS_PWD) : Catch : End Try
    End Sub
    Private Sub TryProtect(ws As SpreadsheetGear.IWorksheet)
        Try : ws.Protect(XLS_PWD) : Catch : End Try
    End Sub
    Private Function D(ws As SpreadsheetGear.IWorksheet, r As Integer, c As Integer) As Decimal
        Dim v = ws.Cells(r, c).Value
        If v Is Nothing OrElse TypeOf v Is DBNull OrElse v.ToString() = "" Then Return 0D
        Return Convert.ToDecimal(v)
    End Function
    Private Sub Recalc_RTLOHB_Forward(ws As SpreadsheetGear.IWorksheet, startCol As Integer)
        Dim r6 As Integer = XLSR("RTLOHB")       'Retailer On Hand (BOM)
        Dim r9 As Integer = XLSR("TYPSLSACTZ")   'Retail Sales (Actualized)
        Dim r13 As Integer = XLSR("TYPGRS")      'Shipments Total (Retail)
        Dim r15 As Integer = XLSR("TYPADJ")      'Returns/Adjustments

        Dim cStart As Integer = Math.Max(startCol, c0 + 1)

        Dim extraMonths As Integer = If(chk12Mos.Checked, 6, 0)
        Dim lastMonthCol As Integer = c0 + 6 + extraMonths

        TryUnprotect(ws)
        Try
            For c As Integer = cStart To lastMonthCol
                Dim p As Integer = c - 1
                Dim addrPrev = ws.Cells(r6, p).Address.Replace("$", "")
                Dim addrPrevAbove = ws.Cells(r6 - 1, p).Address.Replace("$", "")
                Dim addrS = ws.Cells(r9, p).Address.Replace("$", "")
                Dim addrG = ws.Cells(r13, p).Address.Replace("$", "")
                Dim addrA = ws.Cells(r15, p).Address.Replace("$", "")

                If c = c0 + 1 Then
                    ws.Cells(r6, c).Formula =
                    "=IF(" & addrPrev & "="""", " &
                    addrPrevAbove & "-" & addrS & "+" & addrG & "+" & addrA &
                    ", " & addrPrev & "-" & addrS & "+" & addrG & "+" & addrA & ")"
                Else
                    ws.Cells(r6, c).Formula =
                    "=" & addrPrev & "-" & addrS & "+" & addrG & "+" & addrA
                End If
            Next
        Finally
            TryProtect(ws)
        End Try
    End Sub
    Private Sub RefreshTotalsRowsFromHC(rowsToSum() As Integer, colStart As Integer)
        Dim wsTot = workbook.Worksheets(0)
        Dim cStart As Integer = Math.Max(colStart, c0)

        Dim extraMonths As Integer = If(chk12Mos.Checked, 6, 0)
        Dim lastMonthCol As Integer = c0 + 6 + extraMonths
        TryUnprotect(wsTot)
        Try
            For c As Integer = cStart To lastMonthCol
                For Each r In rowsToSum
                    Dim s As Decimal = 0D
                    For Each HC_CODE As String In xls_HC_CODEs
                        Dim ws = workbook.Worksheets(HC_CODE)
                        s += D(ws, r, c)
                    Next
                    wsTot.Cells(r, c).Value = s
                Next
            Next
        Finally
            TryProtect(wsTot)
        End Try
    End Sub
    Private Sub RecalcAfterEdit(ws As SpreadsheetGear.IWorksheet, editedCol As Integer)
        WorkbookView1.GetLock()
        Try
            workbook.WorkbookSet.Calculate()

            Dim wsTotal = workbook.Worksheets(0)
            Dim isTotal As Boolean = Object.ReferenceEquals(ws, wsTotal)
            Dim extraMonths As Integer = If(chk12Mos.Checked, 6, 0)
            Dim lastMonthCol As Integer = c0 + 6 + extraMonths
            If Not isTotal Then
                Recalc_RTLOHB_Forward(ws, editedCol + 1)
            End If
            For c As Integer = Math.Max(editedCol, c0) To lastMonthCol
                WriteTotalRTLOHBFormula(c)
            Next

            workbook.WorkbookSet.Calculate()
        Finally
            WorkbookView1.ReleaseLock()
        End Try
    End Sub
    Private Sub WorkbookView1_CellEndEdit(sender As Object, e As System.EventArgs) _
    Handles WorkbookView1.CellEndEdit
        Dim ac As SpreadsheetGear.IRange = WorkbookView1.ActiveCell
        Dim ws As SpreadsheetGear.IWorksheet = ac.Worksheet
        Dim r As Integer = ac.Row
        Dim c As Integer = ac.Column

        Dim rPlan As Integer = If(XLSR.ContainsKey("TYPSLS"), XLSR("TYPSLS"), -1)
        Dim rAct As Integer = XLSR("TYPSLSACTZ")
        Dim rShip As Integer = XLSR("TYPGRS")
        Dim rAdj As Integer = XLSR("TYPADJ")
        Dim rBOM6 As Integer = XLSR("RTLOHB")

        Dim affects As Boolean =
        (r = rAct OrElse r = rShip OrElse r = rAdj OrElse r = rPlan OrElse r = rBOM6)

        Dim extraMonths As Integer = If(chk12Mos.Checked, 6, 0)
        Dim lastMonthCol As Integer = c0 + 6 + extraMonths
        Dim inMonthRange As Boolean = (c >= c0 AndAlso c <= lastMonthCol)

        If Not (affects AndAlso inMonthRange) Then Exit Sub

        RecalcAfterEdit(ws, c)
    End Sub
    Private Sub WriteTotalRTLOHBFormula(col As Integer)
        Dim wsTot = workbook.Worksheets(0)
        Dim rRTLOHB As Integer = XLSR("RTLOHB")

        TryUnprotect(wsTot)
        Try
            Dim parts As New List(Of String)
            For Each HC_CODE As String In xls_HC_CODEs
                Dim addr = workbook.Worksheets(HC_CODE).Cells(rRTLOHB, col).Address.Replace("$", "")
                parts.Add("'" & HC_CODE & "'!" & addr)
            Next
            wsTot.Cells(rRTLOHB, col).Formula = "=SUM(" & String.Join(",", parts) & ")"
        Finally
            TryProtect(wsTot)
        End Try
    End Sub

End Class