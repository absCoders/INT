Imports Infragistics.Win.UltraWinGrid
Imports System.Text.RegularExpressions
Imports Infragistics.Win.Misc

Public Class RSFRETLI

    Dim filestoImport As New List(Of RetailSalesImporter)
    Dim importNumbersForDelete As New List(Of String)
    Dim editImportNumber As String
    Dim lastClickedCell As UltraGridCell

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        With dst
            ASCMAIN1.sql = "SELECT C1.CUST_CODE,C1.CUST_NAME" & _
                            " FROM " & _
                            " RSTCUSTI CI" & _
                            " JOIN ARTCUST1 C1 ON (CI.CUST_CODE=C1.CUST_CODE)"
            Create_TDA(.Tables.Add, "RSTCUSTI", "**", 0, False, "")

            ASCMAIN1.sql = "SELECT * FROM RSTFILE1"
            Create_TDA(.Tables.Add, "RSTFILE1", "*", 0, True)

            ASCMAIN1.sql = "SELECT RSTIMPR2.*,RSTIMPR1.CUST_CODE,STR.IMPORT_VALUE CUST_STORE_NO_IMPORT,DEPT.IMPORT_VALUE DEPT_CODE_IMPORT, MATL.IMPORT_VALUE MATL_CODE_IMPORT FROM RSTIMPR2 " & _
                            " JOIN RSTIMPR1 ON (RSTIMPR2.IMPORT_NO=RSTIMPR1.IMPORT_NO) " & _
                            " LEFT JOIN RSTIMAP1 STR ON (RSTIMPR1.CUST_CODE=STR.CUST_CODE AND STR.COLUMN_NAME='CUST_STORE_NO' AND STR.RAW_VALUE=RSTIMPR2.CUST_STORE_NO) " & _
                            " LEFT JOIN RSTIMAP1 DEPT ON (RSTIMPR1.CUST_CODE=DEPT.CUST_CODE AND DEPT.COLUMN_NAME='DEPT_CODE' AND DEPT.RAW_VALUE=RSTIMPR2.CUST_DEPT_CODE) " & _
                            " LEFT JOIN RSTIMAP1 MATL ON (RSTIMPR1.CUST_CODE=MATL.CUST_CODE AND MATL.COLUMN_NAME='MATL_CODE' AND MATL.RAW_VALUE=RSTIMPR2.MATL_CODE) " & _
                            " WHERE RSTIMPR2.IMPORT_NO=:PARM1"
            Create_TDA(.Tables.Add, "RSTIMPR2", "**", 0, True, "V")

            ASCMAIN1.sql = "SELECT RSTIMPR1.*,'0' SELECTED FROM RSTIMPR1 WHERE IMPORT_STATUS=:PARM1"
            Create_TDA(.Tables.Add, "RSTIMPR1", "**", 0, True, "V", 0)

            ASCMAIN1.sql = "SELECT RI.*, " & _
                           " CASE WHEN (RC.CUST_OPEN_DOOR_W='1' OR RC.CUST_OPEN_DOOR_M='1') OR RI.REGION_CODE IN ('NEI','SAK') OR (RI.CUST_CODE,RI.CUST_STORE_NO) IN (('FINKS10','ECOMM'),('LUXBOND10','MISC')) THEN '1' ELSE '0' END OPEN_DOOR FROM " & _
                           " RSTRETLI RI LEFT JOIN RSTCUSTS RC ON (RI.CUST_CODE=RC.CUST_CODE AND RI.CUST_STORE_NO=RC.CUST_STORE_NO) " & _
                           " WHERE RI.OPS_YYYYPP=:PARM1 AND RI.REGION_CODE=:PARM2 "
            Create_TDA(.Tables.Add, "RSTRETLI", "**", 0, True, "VV", 4)

            'ASCMAIN1.sql = "SELECT RS.CUST_CODE,C1.CUST_NAME FROM RSTINDYS RS JOIN ARTCUST1 C1 ON (RS.CUST_CODE=C1.CUST_CODE)"
            Create_TDA(.Tables.Add, "RSTINDYS", "*", 0, True, )

            ASCMAIN1.sql = "SELECT * FROM RSTIMPCT"
            Create_TDA(.Tables.Add, "RSTIMPCT", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTIMPCW"
            Create_TDA(.Tables.Add, "RSTIMPCW", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTIMPCM"
            Create_TDA(.Tables.Add, "RSTIMPCM", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTLOADX"
            Create_TDA(.Tables.Add, "RSTLOADX", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTLOADI"
            Create_TDA(.Tables.Add, "RSTLOADI", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTLOADW"
            Create_TDA(.Tables.Add, "RSTLOADW", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM RSTLOADM"
            Create_TDA(.Tables.Add, "RSTLOADM", "**", 0, False)

            ASCMAIN1.sql = "SELECT RSTPLAN1.*,ARTCUST2.CUST_STORE_NAME FROM RSTPLAN1 " & _
                        " LEFT JOIN (SELECT ROW_NUMBER() OVER (PARTITION BY CUST_CODE,LPAD(CUST_STORE_NO,6,'0') ORDER BY CUST_STORE_NO) RNUM,ARTCUST2.* FROM ARTCUST2) ARTCUST2 ON (RSTPLAN1.CUST_CODE=ARTCUST2.CUST_CODE AND RSTPLAN1.CUST_STORE_NO=LPAD(ARTCUST2.CUST_STORE_NO,6,'0') AND ARTCUST2.RNUM=1)" & _
                        " WHERE RSTPLAN1.CUST_CODE=:PARM1 AND RSTPLAN1.OPS_YYYYPP=:PARM2"
            Create_TDA(.Tables.Add, "RSTPLAN1", "**", 0, True, "VV")

            ASCMAIN1.sql = "SELECT * FROM RSTSREG1 " & _
                            " WHERE RSTSREG1.CUST_CODE=:PARM1 ORDER BY RSTSREG1.CUST_REGION_SEQ"
            Create_TDA(.Tables.Add, "RSTSREG1", "**", 0, True, "V")

            ASCMAIN1.sql = "SELECT RSTCUSTS.*,ARTCUST2.CUST_STORE_NAME FROM RSTCUSTS " & _
                            "LEFT JOIN ARTCUST2 ON (RSTCUSTS.CUST_CODE=ARTCUST2.CUST_CODE AND RSTCUSTS.CUST_STORE_NO=ARTCUST2.CUST_STORE_NO)" & _
                            " WHERE RSTCUSTS.CUST_CODE=:PARM1"
            Create_TDA(.Tables.Add, "RSTCUSTS", "**", 0, True, "V")

            ASCMAIN1.sql = "SELECT RSTRETL5.*,ARTCUST2.CUST_STORE_NAME FROM RSTRETL5 " & _
                            " LEFT JOIN (SELECT ROW_NUMBER() OVER (PARTITION BY CUST_CODE,LPAD(CUST_STORE_NO,6,'0') ORDER BY CUST_STORE_NO) RNUM,ARTCUST2.* FROM ARTCUST2) ARTCUST2 ON (RSTRETL5.CUST_CODE=ARTCUST2.CUST_CODE AND RSTRETL5.CUST_STORE_NO=LPAD(ARTCUST2.CUST_STORE_NO,6,'0') AND ARTCUST2.RNUM=1)" & _
                            " WHERE RSTRETL5.CUST_CODE=:PARM1 AND RSTRETL5.OPS_YYYYWW=:PARM2"
            Create_TDA(.Tables.Add, "RSTRETL5", "**", 0, True, "VV")

            ASCMAIN1.sql = "SELECT * FROM RSTPLAN6"
            Create_TDA(.Tables.Add, "RSTPLAN6", "**", 0, True)

            ASCMAIN1.sql = "SELECT * FROM RSTPLANJ WHERE YYYY=:PARM1"
            Create_TDA(.Tables.Add, "RSTPLANJ", "**", 0, True, "N")
        End With

        grdRSTFILE1.DataSource = dst.Tables("RSTFILE1")
        grdRSTIMPR2.DataSource = dst.Tables("RSTIMPR2")
        grdRSTRETL5.DataSource = dst.Tables("RSTRETL5")
        grdRSTSREG1.DataSource = dst.Tables("RSTSREG1")
        grdRSTCUSTS.DataSource = dst.Tables("RSTCUSTS")
        grdRSTPLANJ.DataSource = dst.Tables("RSTPLANJ")

        Sort_grdColumns(grdRSTIMPR1, "import_no")
        Sort_grdColumns(grdRSTIMPR2, "CUST_STORE_NO")

        grdRSTIMPR1.DataSource = dst.Tables("RSTIMPR1")

        grdRSTINDYS.DataSource = dst.Tables("RSTINDYS")

        grdRSTRETLI.DataSource = dst.Tables("RSTRETLI")
        grdRSTRETLI.DisplayLayout.Bands(0).ColumnFilters("OPEN_DOOR").FilterConditions.Add(FilterComparisionOperator.Equals, "1") 'only show open doors by default

        grdRSTPLAN1.DataSource = dst.Tables("RSTPLAN1")
        grdRSTPLAN6.DataSource = dst.Tables("RSTPLAN6")

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP

        Create_Summary(grdRSTIMPR2, "CUST_STORE_NO", "Count")
        Create_Summary(grdRSTIMPR2, "AMT_SOLD", "Sum")
        Create_Summary(grdRSTIMPR2, "AMT_SOLD_MTD", "Sum")
        Create_Summary(grdRSTIMPR2, "AMT_SOLD_STD", "Sum")
        Create_Summary(grdRSTIMPR2, "AMT_PLAN", "Sum")

        Create_Summary(grdRSTRETL5, "AMT_SOLD", "Sum")

        Create_Summary(grdRSTPLAN1, "AMT_PLAN", "Sum")

        Create_Summary(grdRSTRETLI, "AMT_SOLD", "Sum")
        Create_Summary(grdRSTRETLI, "AMT_SOLD_LY", "Sum")

        InitializeImportCountGrid(grdRSTIMPCT)
        grdRSTIMPCT.DataSource = dst.Tables("RSTIMPCT")

        InitializeImportCountGrid(grdRSTLOADX)
        grdRSTLOADX.DataSource = dst.Tables("RSTLOADX")

        InitializeIndependentCountGrid(grdRSTLOADI)
        grdRSTLOADI.DataSource = dst.Tables("RSTLOADI")

        InitializeComboBoxes()

        Set_cmbYP("RSRYP", ASCMAIN1.CYP, -24, 0, -1)
        Set_cmbYW("RSRYW", ASCMAIN1.CYW, -100, 0, -1)

        cmbPlanYear.DataSource = New Integer() {2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025}
        cmbPlanYear.SelectedText = Now.AddMonths(5).Year

        InitializeWeekDropDown()
        LoadImportFileHeaders(ImportStatus.NotLoaded)
    End Sub


    Private Sub InitializeWeekDropDown()
        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT Distinct TO_CHAR(TO_DATE(MM,'MM'),'Mon') || ' Week ' || REL_WEEK || ' ' || YM WEEK_DESC,YM,MM,REL_WEEK, YM || ',' || MM || ',' || REL_WEEK WEEK_STRING FROM RSTCLND1 WHERE YYYYWW >= :PARM1 and YYYYWW <= :PARM2 order by YM,mm,rel_week", "", "VV", New Object() {ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -100), ASCMAIN1.CYW})

        UltraCombo1.SetDataBinding(dt, Nothing)
        UltraCombo1.ValueMember = "WEEK_STRING"
        UltraCombo1.DisplayMember = "WEEK_DESC"

        UltraCombo1.SelectedRow = UltraCombo1.Rows(UltraCombo1.Rows.Count - 2)

        UltraCombo1.DisplayLayout.Bands(0).Columns("YM").Hidden = True
        UltraCombo1.DisplayLayout.Bands(0).Columns("MM").Hidden = True
        UltraCombo1.DisplayLayout.Bands(0).Columns("REL_WEEK").Hidden = True
        UltraCombo1.DisplayLayout.Bands(0).Columns("WEEK_STRING").Hidden = True
        UltraCombo1.DisplayLayout.Bands(0).Columns("WEEK_DESC").Header.Caption = "Retail Week"
    End Sub

    Private Function GetYYYYWWFromDropDown(ByVal week_string As String, ByVal custCode As String) As String
        Dim stringParts As String() = week_string.Split(",")
        Return ASCDATA1.GetDataValue("SELECT YYYYWW FROM RSTCLND1 WHERE YM=:PARM1 AND MM=:PARM2 AND REL_WEEK=:PARM3 AND CUST_CODE=:PARM4", "VVVV", New Object() {stringParts(0), stringParts(1), stringParts(2), If(custCode = "ECOMSALE10", "BLOOMIES10", custCode)})
    End Function

    Private Sub InitializeComboBoxes()
        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT CUST_CODE,CUST_NAME FROM ARTCUST1 WHERE CUST_CODE IN ('BLOOMIES10','NORDSTR10','NEIMANM10','SAKSFIF10','HOLTREN10','ECOMSALE10')")

        cmbPlanCustomers.SetDataBinding(dt, Nothing)
        cmbPlanCustomers.ValueMember = "CUST_CODE"
        cmbPlanCustomers.DisplayMember = "CUST_NAME"

        cmbPlanCustomers.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbPlanCustomers.DataSource = dt

        cmbCustomersNationalEntry.SetDataBinding(dt, Nothing)
        cmbCustomersNationalEntry.ValueMember = "CUST_CODE"
        cmbCustomersNationalEntry.DisplayMember = "CUST_NAME"

        cmbCustomersNationalEntry.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbCustomersNationalEntry.DataSource = dt

        Dim dt3 As DataTable = ASCDATA1.GetDataTable("SELECT REGION_CODE,REGION_DESC FROM SOTSREG1 WHERE REGION_CODE IN (SELECT REGION_CODE FROM RSTIREG1)")

        cmbIndRegion.SetDataBinding(dt3, Nothing)
        cmbIndRegion.ValueMember = "REGION_CODE"
        cmbIndRegion.DisplayMember = "REGION_DESC"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Import Raw Files"
                Dim importFileNames As String() = GetImportFileNames()
                If importFileNames Is Nothing Then Exit Sub

                ASCMAIN1.Progress("Loading and verifying files...")
                filestoImport.Clear()
                For Each importFileName As String In importFileNames
                    Dim fileImporter As RetailSalesImporter
                    Try
                        fileImporter = RetailSalesImporter.CreateImporter(importFileName)
                        If fileImporter IsNot Nothing Then
                            filestoImport.Add(fileImporter)
                        Else
                            Dim x = 0
                        End If
                    Catch ex As Exception
                        If MessageBox.Show(String.Format("Error: {0}Continue importing?", ex.Message & vbCrLf & vbCrLf), "Error", MessageBoxButtons.YesNo) = DialogResult.No Then
                            ASCMAIN1.Progress("")
                            Exit Sub
                        End If
                    End Try
                Next
                ASCMAIN1.Progress("")

            Case "Load Selected Data", "Retract Imported Data"
                If dst.Tables("RSTIMPR1").Select("SELECTED='1'").Length = 0 Then
                    EMsg &= "No rows selected"
                End If

                If eItemKey = "Load Selected Data" Then
                    Dim importsNeedingMapping As String = ""
                    Dim delim As String = ""
                    For Each dr As DataRow In dst.Tables("RSTIMPR1").Select("SELECTED='1'")
                        Fill_Records("RSTIMPR2", New Object() {dr.Item("IMPORT_NO")})
                        If dst.Tables("RSTIMPR2").AsEnumerable().Count(Function(row) (row.Item("DEPT_CODE_IMPORT") & "" = "" Or row.Item("MATL_CODE_IMPORT") & "" = "") Or (row.Item("CUST_STORE_NO_IMPORT") & "" = "" And Not Regex.IsMatch(row.Item("CUST_STORE_NO"), "(\d{1,6}|DIRECT)"))) > 0 Then
                            importsNeedingMapping &= delim & dr.Item("IMPORT_NO")
                            delim = ","
                        End If
                    Next

                    If importsNeedingMapping.Length > 0 Then
                        EMsg &= "You must create mappings for the following imports: " & importsNeedingMapping & vbCrLf
                    End If
                End If

            Case "Load"
                Select Case EntryMode
                    Case "L"
                        Validate_Code("OPS_YYYYPP")
                        'Validate_Code("REGION_CODE")
                    Case "N"

                    Case "P"
                        If optPlanType.Value = "N" Then
                            If cmbPlanCustomers.Value = "" Then
                                EMsg &= "No customer selected"
                            End If
                        End If
                End Select

            Case "Update"
                Select Case EntryMode
                    Case "P"
                        If optPlanType.Value = "N" Then
                            If optPlanPeriodSeason.Value = "P" Then
                                If dst.Tables("RSTPLAN1").AsEnumerable().Any(Function(row) row.RowState <> DataRowState.Deleted AndAlso (row.Item("MATL_CODE") & "" = "" Or row.Item("DEPT_CODE") & "" = "")) Then
                                    EMsg &= "There are rows with invalid department or material codes"
                                End If
                            Else
                                If dst.Tables("RSTPLAN6").AsEnumerable().Any(Function(row) row.RowState <> DataRowState.Deleted AndAlso (row.Item("MATL_CODE") & "" = "" Or row.Item("DEPT_CODE") & "" = "")) Then
                                    EMsg &= "There are rows with invalid department or material codes"
                                End If
                            End If
                        Else

                        End If
                End Select

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Import Raw Files"
                ASCMAIN1.Progress("Loading file data...")
                For Each fileToImport In filestoImport
                    If fileToImport Is Nothing OrElse Not fileToImport.IsValid Then
                        MsgBox(String.Format("File ({0}) is not valid for customer: {1} ", fileToImport.FileName, fileToImport.CustomerCode))
                        Continue For
                    End If

                    Try
                        If fileToImport.retailImportType = RetailImportType.SaksConsignmentExcel Then
                            Using wsPicker As New ASFMSGBF()
                                Dim dt As New DataTable()
                                dt = CType(fileToImport, SaksConsignmentImportExcel).GetWorksheets()

                                Dim wsName As String =
                                                wsPicker.Get_cmb_from_User("Worksheet:",
                                                String.Format("Select Worksheet for {0}", fileToImport.FileName), dt)

                                CType(fileToImport, SaksConsignmentImportExcel).Import(wsName)
                            End Using
                        Else
                            fileToImport.Import()
                        End If


                        If fileToImport.SalesWeek = "" Then
                            EMsg &= String.Format("Unable to determine time frame of selected file {0}", fileToImport.FileName)
                        End If

                        ImportRawFileDataToDB(fileToImport)
                    Catch ex As Exception
                        MsgBox(String.Format("Error importing file {0}:" & vbCrLf & vbCrLf & "{1}", fileToImport.FileName, ex.Message))
                    End Try
                Next

                RefreshImportCountGrid()
                LoadImportFileHeaders(ImportStatus.NotLoaded)
                optMode.CheckedIndex = 1
                ASCMAIN1.Progress("")

            Case "Load Selected Data"
                CreateImportedData()
                LoadImportFileHeaders(ImportStatus.NotLoaded)

            Case "Retract Imported Data"
                RetractImportedData()
                LoadImportFileHeaders(ImportStatus.LoadedAlready)

            Case "Restore Deleted"
                RestoreDeletedData()
                LoadImportFileHeaders(ImportStatus.DeletedFromQueue)

            Case "Load"
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

        Select Case optMode.Value
            Case Mode.IndependentEntry
                EntryMode = "L"
            Case Mode.NationalEntry
                EntryMode = "N"
            Case Mode.PlanEntry
                EntryMode = "P"
            Case Mode.StoreMaintenance
                EntryMode = "M"
            Case Mode.ReportSettings
                EntryMode = "R"
        End Select

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Raw Files").Visible = (optMode.Value = Mode.RawFiles)
                .Groups("Imported Data").Visible = (optMode.Value = Mode.ImportedData)

                .Groups("Import Counts").Visible = (optMode.Value = Mode.ImportedData) Or (optMode.Value = Mode.RawFiles)

                .Groups("Batch Control").Visible = (optMode.Value = Mode.RawFiles Or optMode.Value = Mode.ImportedData)
                .Groups("Batch Control").Items("Import Raw Files").Visible = (optMode.Value = Mode.RawFiles)
                .Groups("Batch Control").Items("Load Selected Data").Visible = (optMode.Value = Mode.ImportedData And optImportData.Value = ImportStatus.NotLoaded)
                .Groups("Batch Control").Items("Retract Imported Data").Visible = (optMode.Value = Mode.ImportedData And optImportData.Value = ImportStatus.LoadedAlready)
                .Groups("Batch Control").Items("Restore Deleted").Visible = (optMode.Value = Mode.ImportedData And optImportData.Value = ImportStatus.DeletedFromQueue)

                .Groups("Entry Options").Visible = (optMode.Value = Mode.NationalEntry)

                .Groups("Independent Entry").Visible = (optMode.Value = Mode.IndependentEntry)

                .Groups("Screen Control").Visible = (optMode.Value = Mode.IndependentEntry Or optMode.Value = Mode.NationalEntry Or optMode.Value = Mode.PlanEntry Or optMode.Value = Mode.StoreMaintenance)
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = If(ScreenMode And (EntryMode = "P" Or EntryMode = "L" Or EntryMode = "N" Or EntryMode = "M"), DefaultableBoolean.True, DefaultableBoolean.False)
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With

            With tabImportMode
                If .Tabs.Count > 0 Then
                    .Tabs(Mode.RawFiles).Visible = (optMode.Value = Mode.RawFiles)
                    .Tabs(Mode.ImportedData).Visible = (optMode.Value = Mode.ImportedData)
                    .Tabs(Mode.NationalEntry).Visible = (optMode.Value = Mode.NationalEntry)
                    .Tabs(Mode.IndependentEntry).Visible = (optMode.Value = Mode.IndependentEntry)
                    .Tabs(Mode.PlanEntry).Visible = (optMode.Value = Mode.PlanEntry)
                    .Tabs(Mode.StoreMaintenance).Visible = (optMode.Value = Mode.StoreMaintenance)
                    .Tabs(Mode.ReportSettings).Visible = (optMode.Value = Mode.ReportSettings)

                    .SelectedTab = .Tabs(optMode.Value)
                End If
            End With

            optMode.Enabled = Not ScreenMode
            grdRSTRETL5.Visible = (optMode.Value = Mode.NationalEntry) And ScreenMode = True
            grdRSTRETLI.Visible = (optMode.Value = Mode.IndependentEntry) And ScreenMode = True
            grdRSTPLAN1.Visible = (optMode.Value = Mode.PlanEntry) And ScreenMode = True
            grdRSTPLAN6.Visible = (optMode.Value = Mode.PlanEntry) And ScreenMode = True
            grdRSTPLANJ.Visible = (optMode.Value = Mode.PlanEntry) And ScreenMode = True
            grdRSTCUSTS.Visible = (optMode.Value = Mode.StoreMaintenance) And ScreenMode = True
            grdRSTSREG1.Visible = (optMode.Value = Mode.StoreMaintenance) And ScreenMode = True
            SplitContainer8.Panel1Collapsed = (optMode.Value = Mode.IndependentEntry) And ScreenMode
            SplitContainer8.Panel2Collapsed = (optMode.Value = Mode.IndependentEntry) And Not ScreenMode

            Set_Read_Only(UltraGroupBox6, (optMode.Value = Mode.NationalEntry) And ScreenMode = True)
            Set_Read_Only(UltraGroupBox4, (optMode.Value = Mode.IndependentEntry) And ScreenMode = True)
            Set_Read_Only(UltraGroupBox5, (optMode.Value = Mode.PlanEntry) And ScreenMode = True)

            If grdRSTRETL5.DisplayLayout.Bands(0).Columns.Count > 0 Then
                grdRSTRETL5.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Hidden = (optMode.Value = Mode.NationalEntry) And ScreenMode = True And cmbCustomersNationalEntry.Value = "ECOMSALE10"
                grdRSTRETL5.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Hidden = (optMode.Value = Mode.NationalEntry) And ScreenMode = True And cmbCustomersNationalEntry.Value = "ECOMSALE10"
            End If

            splPlanEntry.Panel1Collapsed = (optPlanPeriodSeason.Value = "S")
            splPlanEntry.Panel2Collapsed = (optPlanPeriodSeason.Value = "P")

            SplitContainer9.Panel1Collapsed = (optPlanType.Value = "N")
            SplitContainer9.Panel2Collapsed = (optPlanType.Value = "J")

        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        If dst.Tables.Count > 0 Then
            dst.Tables("RSTIMPR1").Clear()
            dst.Tables("RSTIMPR2").Clear()
            dst.Tables("RSTCUSTS").Clear()
            dst.Tables("RSTSREG1").Clear()
            dst.Tables("RSTRETL5").Clear()
            dst.Tables("RSTPLAN1").Clear()
            dst.Tables("RSTPLAN6").Clear()
        End If
        dst.EnforceConstraints = True
    End Sub

    Private Function GetSeasonStartPeriod(ByVal seasonInput As String)
        Dim seasonStartPeriod As String = seasonInput.Substring(0, 4) & "01"
        If Val(seasonInput.Substring(4)) >= 7 Then
            seasonStartPeriod = seasonInput.Substring(0, 4) & "07"
        End If
        Return seasonStartPeriod
    End Function

    Sub FillRSTPLAN6(ByVal custCode As String, ByVal periodInSeason As String)

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE RSTPLAN6")
        Dim sql As String = " INSERT INTO RSTPLAN6" &
                            " SELECT CUST_CODE,CUST_STORE_NO,DEPT_CODE,MATL_CODE, " &
                            " SUM(CASE WHEN OPS_YYYYPP=:PARM2 THEN AMT_PLAN ELSE 0 END) AMT_PLAN_01, " &
                            " SUM(CASE WHEN OPS_YYYYPP=PERIOD_CALC(:PARM2,1) THEN AMT_PLAN ELSE 0 END) AMT_PLAN_02, " &
                            " SUM(CASE WHEN OPS_YYYYPP=PERIOD_CALC(:PARM2,2) THEN AMT_PLAN ELSE 0 END) AMT_PLAN_03, " &
                            " SUM(CASE WHEN OPS_YYYYPP=PERIOD_CALC(:PARM2,3) THEN AMT_PLAN ELSE 0 END) AMT_PLAN_04, " &
                            " SUM(CASE WHEN OPS_YYYYPP=PERIOD_CALC(:PARM2,4) THEN AMT_PLAN ELSE 0 END) AMT_PLAN_05, " &
                            " SUM(CASE WHEN OPS_YYYYPP=PERIOD_CALC(:PARM2,5) THEN AMT_PLAN ELSE 0 END) AMT_PLAN_06 " &
                            "        FROM " &
                            "        RSTPLAN1 " &
                            " WHERE CUST_CODE=:PARM1 AND OPS_YYYYPP BETWEEN :PARM2 AND PERIOD_CALC(:PARM2,5)" &
                            " group by CUST_CODE,CUST_STORE_NO,DEPT_CODE,MATL_CODE"
        Dim startPeriod = GetSeasonStartPeriod(periodInSeason)
        For i = 1 To 6
            Dim curPeriod = ASCMAIN1.Period_Calc(startPeriod, i - 1)
            Dim curDate = New Date(Convert.ToInt32(curPeriod.Substring(0, 4)), Convert.ToInt32(curPeriod.Substring(4, 2)), 15)
            curDate = curDate.AddMonths(7)
            grdRSTPLAN6.DisplayLayout.Bands(0).Columns("AMT_PLAN_" & i.ToString("00")).Header.Caption = curDate.ToString("MMM")
        Next

        ASCDATA1.ExecuteSQL(sql, "VV", New Object() {custCode, startPeriod})
    End Sub

    Private Sub FillRSTPLAN1FromRSTPLAN6(custCode As String, period As String)
        Dim SQL As String = "DECLARE" _
                        & " PERIOD_START VARCHAR2(6) := :PARM1;" _
                        & " BEGIN" _
                        & " DELETE FROM RSTPLAN1 WHERE CUST_CODE=:PARM2 AND OPS_YYYYPP BETWEEN :PARM1 AND PERIOD_CALC(:PARM1,5);" _
                        & " FOR I IN 1..6 LOOP" _
                        & " EXECUTE IMMEDIATE " _
                        & " 'INSERT INTO RSTPLAN1 (CUST_CODE,CUST_STORE_NO,DEPT_CODE,MATL_CODE,OPS_YYYYPP,AMT_PLAN)" _
                        & "  SELECT CUST_CODE,CUST_STORE_NO,DEPT_CODE,MATL_CODE,PERIOD_CALC(''' || PERIOD_START || ''',' || I || '-1),AMT_PLAN_' || TO_CHAR(I,'FM00') ||" _
                        & " ' FROM RSTPLAN6';" _
                        & " END LOOP;" _
                        & " END;"
        ASCDATA1.ExecuteSQL(SQL, "VV", New Object() {GetSeasonStartPeriod(period), custCode})
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        If EntryMode = "L" Then
            BindIndependentsDropDowns(cmbIndRegion.Value)
            Fill_Records("RSTRETLI", New Object() {cmbIndependentPeriod.Value, cmbIndRegion.Value})
            Dim sqlNewCustomers As String = "SELECT DISTINCT NVL(SR.SREP_CODE,'NA') REGION_CODE,R5.CUST_CODE,R5.CUST_STORE_NO,NVL(RC.CUST_OPEN_DOOR_W,'1') OPEN_DOOR FROM RSTRETL5 R5 " & _
                            " JOIN ARTCUST1 C1 ON (R5.CUST_CODE=C1.CUST_CODE) " & _
                            " LEFT JOIN SOTSREP1 SR ON (C1.SREP_CODE=SR.SREP_CODE) " & _
                            " LEFT JOIN RSTCUSTS RC ON (R5.CUST_CODE=RC.CUST_CODE AND R5.CUST_STORE_NO=RC.CUST_STORE_NO) WHERE " & _
                            " (NVL(C1.TRADE_CLASS_CODE,'IND')='IND' OR C1.CUST_CODE IN ('NMLASTCALL','SAKSOFF5TH')) " & _
                            " AND R5.OPS_YYYYPP = :PARM1 AND NVL(SR.SREP_CODE,'NA')=:PARM2" & _
                            " AND NOT EXISTS (SELECT 1 FROM RSTRETL5 WHERE CUST_CODE=R5.CUST_CODE AND CUST_STORE_NO=R5.CUST_STORE_NO AND OPS_YYYYPP=:PARM3)"

            Dim dtNewCusts As DataTable = ASCDATA1.GetDataTable(sqlNewCustomers, "newCusts", "VVV", New Object() {ASCMAIN1.Period_Calc(cmbIndependentPeriod.Value, -1), cmbIndRegion.Value, ASCMAIN1.Period_Calc(cmbIndependentPeriod.Value, -12)})

            For Each row As DataRow In dtNewCusts.Rows
                dst.Tables("RSTRETLI").Rows.Add(New Object() {cmbIndependentPeriod.Value, row.Item("REGION_CODE"), row.Item("CUST_CODE"), row.Item("CUST_STORE_NO"), "WM", "S", Nothing, Nothing, row.Item("OPEN_DOOR")})
            Next

        ElseIf EntryMode = "P" Then
            If optPlanType.Value = "N" Then
                If optPlanPeriodSeason.Value = "P" Then
                    Fill_Records("RSTPLAN1", New Object() {cmbPlanCustomers.Value, Absx1.txtFor("OPS_YYYYPP").Text})
                Else
                    FillRSTPLAN6(cmbPlanCustomers.Value, Absx1.txtFor("OPS_YYYYPP").Text)
                    Fill_Records("RSTPLAN6")
                End If
            Else
                BindJHPlanCustomerNames()
                Fill_Records("RSTPLANJ", New Object() {cmbPlanYear.Value})
            End If
        ElseIf EntryMode = "N" Then
            BindNationalStoreNames(cmbCustomersNationalEntry.Value)
            editImportNumber = ASCMAIN1.Next_Control_No("RSTIMPR1.IMPORT_NO")
            Fill_Records("RSTRETL5", New Object() {cmbCustomersNationalEntry.Value, GetYYYYWWFromDropDown(UltraCombo1.Value, cmbCustomersNationalEntry.Value)})
        ElseIf EntryMode = "M" Then
            BindStoreMaintenanceDropDown(txtCustomersSM.Value)
            Fill_Records("RSTSREG1", New Object() {txtCustomersSM.Value})
            Fill_Records("RSTCUSTS", New Object() {txtCustomersSM.Value})
        End If
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        If EntryMode = "L" Then
            Update_Record_TDA("RSTRETLI")
        ElseIf EntryMode = "P" Then
            If optPlanType.Value = "N" Then
                If optPlanPeriodSeason.Value = "P" Then
                    Update_Record_TDA("RSTPLAN1")
                Else
                    Update_Record_TDA("RSTPLAN6")
                    FillRSTPLAN1FromRSTPLAN6(cmbPlanCustomers.Value, Absx1.txtFor("OPS_YYYYPP").Text)
                End If
            Else
                Update_Record_TDA("RSTPLANJ")
            End If
        ElseIf EntryMode = "N" Then

            For Each row As DataRow In dst.Tables("RSTRETL5").Select("", "", DataViewRowState.Deleted)
                For Each rowAdded In dst.Tables("RSTRETL5").Select(String.Format("CUST_CODE='{0}' AND CUST_STORE_NO='{1}' AND DEPT_CODE='{2}' AND MATL_CODE='{3}' AND OPS_YYYYWW='{4}'", row.Item("CUST_CODE", DataRowVersion.Original), row.Item("CUST_STORE_NO", DataRowVersion.Original), row.Item("DEPT_CODE", DataRowVersion.Original), row.Item("MATL_CODE", DataRowVersion.Original), row.Item("OPS_YYYYWW", DataRowVersion.Original)), "", DataViewRowState.Added)
                    row.RejectChanges()
                    row.SetModified()
                    row.Item("AMT_SOLD") = rowAdded.Item("AMT_SOLD")
                    dst.Tables("RSTRETL5").Rows.Remove(rowAdded)
                    Exit For
                Next
            Next

            For Each row As DataRow In dst.Tables("RSTRETL5").Rows
                If row.RowState <> DataRowState.Unchanged Then
                    ASCDATA1.ExecuteSQL("INSERT INTO RSTAUDT1 (IMPORT_NO,CUST_CODE,CUST_STORE_NO,DEPT_CODE,MATL_CODE,OPS_YYYYWW,EDIT_DATE,OLD_VALUE,NEW_VALUE) VALUES (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,:PARM7,:PARM8)", _
                                        "VVVVVVNN", New Object() {editImportNumber, cmbCustomersNationalEntry.Value,
                                                                   If(row.RowState <> DataRowState.Deleted, row.Item("CUST_STORE_NO", DataRowVersion.Default), row.Item("CUST_STORE_NO", DataRowVersion.Original)), _
                                                                   If(row.RowState <> DataRowState.Deleted, row.Item("DEPT_CODE", DataRowVersion.Default), row.Item("DEPT_CODE", DataRowVersion.Original)),
                                                                   If(row.RowState <> DataRowState.Deleted, row.Item("MATL_CODE", DataRowVersion.Default), row.Item("MATL_CODE", DataRowVersion.Original)), GetYYYYWWFromDropDown(Absx1.cmbFor("RSRYW").Value, cmbCustomersNationalEntry.Value), _
                                                                   If(row.RowState <> DataRowState.Added, row.Item("AMT_SOLD", DataRowVersion.Original), Nothing), _
                                                                   If(row.RowState <> DataRowState.Deleted, row.Item("AMT_SOLD", DataRowVersion.Current), Nothing)})
                End If
            Next
            Update_Record_TDA("RSTRETL5")
        ElseIf EntryMode = "M" Then
            Update_Record_TDA("RSTSREG1")
            Update_Record_TDA("RSTCUSTS")
        End If
        CommitTrans()
    End Sub

    Private Sub LoadImportFileHeaders(ByVal importStatus As String)
        If dst.Tables.Contains("RSTIMPR1") Then
            dst.Tables("RSTIMPR1").Clear()
            dst.Tables("RSTIMPR2").Clear()
            grdRSTIMPR2.Text = "Raw File Data"
            Fill_Records("RSTIMPR1", importStatus)
        End If
    End Sub
    Private Sub LoadImportFileDetails(ByVal importNo As String)
        If dst.Tables.Contains("RSTIMPR2") Then
            dst.Tables("RSTIMPR2").Clear()
            Fill_Records("RSTIMPR2", importNo)
        End If
    End Sub
    Private Sub LoadRawFileHeaders()
        If dst.Tables.Contains("RSTFILE1") Then
            dst.Tables("RSTFILE1").Clear()
            Fill_Records("RSTFILE1")
        End If
    End Sub

    Function ImportRawFileDataToDB(ByVal importFile As RetailSalesImporter) As String
        BeginTrans()

        'There should only ever be multiple weeks when importing a Holt Renfrew file
        Dim weeksToImport As List(Of String) = importFile.ImportedRetailData.AsEnumerable.Select(Function(row) row.Item("OPS_YYYYWW").ToString()).Distinct().ToList()
        Dim importControlNumber As String = Nothing
        Dim updateTime = Now

        For Each week As String In weeksToImport
            Dim curWeek = week
            importControlNumber = ASCMAIN1.Next_Control_No("RSTIMPR1.IMPORT_NO")

            Dim period As String = importFile.ImportedRetailData.AsEnumerable.Where(Function(row) row.Item("OPS_YYYYWW") = curWeek).Select(Function(row) row.Item("OPS_YYYYPP").ToString).First()

            Dim drRSTIMPR1 As DataRow = dst.Tables("RSTIMPR1").NewRow
            drRSTIMPR1.Item("IMPORT_NO") = importControlNumber
            drRSTIMPR1.Item("CUST_CODE") = importFile.CustomerCode
            drRSTIMPR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            drRSTIMPR1.Item("INIT_DATE") = updateTime
            drRSTIMPR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            drRSTIMPR1.Item("LAST_DATE") = updateTime
            drRSTIMPR1.Item("IMPORT_STATUS") = "0"
            drRSTIMPR1.Item("OPS_YYYYWW") = week
            drRSTIMPR1.Item("OPS_YYYYPP") = period
            drRSTIMPR1.Item("IMPORT_FILENAME") = importFile.FileName
            dst.Tables("RSTIMPR1").Rows.Add(drRSTIMPR1)

            For Each drRSTRETLI As DataRow In importFile.ImportedRetailData.Select(String.Format("OPS_YYYYWW='{0}'", week))
                Dim drRSTIMPR2 As DataRow = dst.Tables("RSTIMPR2").NewRow
                drRSTIMPR2.Item("IMPORT_NO") = importControlNumber
                drRSTIMPR2.Item("CUST_STORE_NO") = drRSTRETLI.Item("CUST_STORE_NO")
                drRSTIMPR2.Item("CUST_DEPT_CODE") = drRSTRETLI.Item("CUST_DEPT_CODE")
                drRSTIMPR2.Item("AMT_SOLD") = drRSTRETLI.Item("AMT_SOLD")
                drRSTIMPR2.Item("AMT_SOLD_MTD") = drRSTRETLI.Item("AMT_SOLD_MTD")
                drRSTIMPR2.Item("AMT_SOLD_STD") = drRSTRETLI.Item("AMT_SOLD_STD")
                drRSTIMPR2.Item("AMT_PLAN") = drRSTRETLI.Item("AMT_PLAN")
                drRSTIMPR2.Item("MEMO") = drRSTRETLI.Item("MEMO")
                drRSTIMPR2.Item("MATL_CODE") = drRSTRETLI.Item("MATL_CODE")
                dst.Tables("RSTIMPR2").Rows.Add(drRSTIMPR2)
            Next
        Next

        If importFile.retailImportType <> RetailImportType.SaksConsignmentExcel Then
            Dim drRSTFILE1 As DataRow = dst.Tables("RSTFILE1").NewRow()
            drRSTFILE1.Item("IMPORT_FILENAME") = importFile.FileName
            drRSTFILE1.Item("IMPORT_FILESIZE") = importFile.FileSize
            drRSTFILE1.Item("IMPORT_DATETIME") = updateTime
            drRSTFILE1.Item("IMPORT_NO") = importControlNumber
            dst.Tables("RSTFILE1").Rows.Add(drRSTFILE1)
        End If

        Try
            Update_Record_TDA("RSTIMPR1")
            Update_Record_TDA("RSTIMPR2")
            Update_Record_TDA("RSTFILE1")
            CommitTrans("")
        Catch ex As Exception
            Rollback(String.Format("Import failed: {0}", ex.Message))
            Return ""
        End Try

        Return importControlNumber
    End Function

    Sub CreateImportedData()
        For Each drRSTIMPR1 As DataRow In dst.Tables("RSTIMPR1").Select("SELECTED='1'", "OPS_YYYYWW")
            Dim importNo As String = drRSTIMPR1.Item("IMPORT_NO")

            Try
                BeginTrans()
                ASCDATA1.ExecuteSP("RETAIL.CREATE_AUDITED_RETAIL", "V", New Object() {importNo}, New String() {"IMPORT_NO"})
                CommitTrans("")
            Catch ex As Exception
                Rollback(String.Format("Error importing: {0}", ex.Message))
            End Try
        Next
    End Sub


    Private Sub RetractImportedData()
        For Each drRSTIMPR1 As DataRow In dst.Tables("RSTIMPR1").Select("SELECTED='1'")
            Dim importNo As String = drRSTIMPR1.Item("IMPORT_NO")
            Try
                BeginTrans()
                Dim retval As String = ASCDATA1.ExecuteSP("RETAIL.RETRACT_AUDITED_RETAIL", "V", New Object() {importNo}, New String() {"IMPORT_NO"})
                CommitTrans()
            Catch ex As Exception
                Rollback(String.Format("Error retracting: {0}", ex.Message))
            End Try
        Next
    End Sub

    Private Sub RestoreDeletedData()
        For Each drRSTIMPR1 As DataRow In dst.Tables("RSTIMPR1").Select("SELECTED='1'")
            Dim importNo As String = drRSTIMPR1.Item("IMPORT_NO")
            Try
                BeginTrans()
                Dim retval As String = ASCDATA1.ExecuteSQL("UPDATE RSTIMPR1 SET IMPORT_STATUS='0' WHERE IMPORT_NO=:PARM1 AND IMPORT_STATUS='D'", "V", New Object() {importNo})
                CommitTrans()
            Catch ex As Exception
                Rollback(String.Format("Error restoring: {0}", ex.Message))
            End Try
        Next
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTIMPR1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTIMPR2, "SSPB", "Show Filter", "Show GroupBox", "Set Mapping")
        Load_Popup_Menu(grdRSTPLAN1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTRETL5, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTRETLI, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdRSTCUSTS, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdRSTBUDR1"
                '    If grdRSTRETLA.Tag = "" Then
                '        e.Cancel = True
                '    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"

            Case "Show GroupBox"
                'If grd IsNot Nothing Then
                '    Dim tlb_sbt As StateButtonTool = DirectCast(e.Tool, StateButtonTool)
                '    grd.DisplayLayout.Bands(0).ColHeadersVisible = tlb_sbt.Checked
                'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Set Mapping"
                Dim custCode = lastClickedCell.Row.Cells("CUST_CODE").Value
                Dim impNo As String = lastClickedCell.Row.Cells("IMPORT_NO").Value
                Select Case lastClickedCell.Column.Key
                    Case "CUST_STORE_NO", "CUST_STORE_NO_IMPORT"
                        Dim input As New ASFMSGBF()
                        Dim rawValue = lastClickedCell.Row.Cells("CUST_STORE_NO").Value
                        Dim storeMap As String = input.Get_txt_from_User(String.Format("Input store no for {0}:", rawValue), "Store No", False, 6)
                        If storeMap <> "" Then
                            storeMap = storeMap.PadLeft(6, "0").ToUpper()
                            If Not Regex.IsMatch(storeMap, "(\d{1,6}|DIRECT)") Then
                                MsgBox("You entered an invalid store no")
                            Else
                                ASCDATA1.ExecuteSQL(
                                    "MERGE INTO RSTIMAP1 RM " &
                                    "USING (SELECT :PARM1 CUST_CODE, :PARM2 COLUMN_NAME, :PARM3 RAW_VALUE, :PARM4 IMPORT_VALUE FROM DUAL) VALS " &
                                    "ON (RM.CUST_CODE=VALS.CUST_CODE AND RM.COLUMN_NAME=VALS.COLUMN_NAME AND RM.RAW_VALUE=VALS.RAW_VALUE) " &
                                    "WHEN MATCHED THEN " &
                                    "UPDATE SET RM.IMPORT_VALUE=VALS.IMPORT_VALUE " &
                                    "WHEN NOT MATCHED THEN " &
                                    "INSERT (CUST_CODE,COLUMN_NAME,RAW_VALUE,IMPORT_VALUE) VALUES (VALS.CUST_CODE,VALS.COLUMN_NAME,VALS.RAW_VALUE,VALS.IMPORT_VALUE)",
                                    "VVVV", New Object() {custCode, "CUST_STORE_NO", rawValue, storeMap})
                            End If
                        End If

                    Case "CUST_DEPT_CODE", "DEPT_CODE_IMPORT"
                        Dim input As New ASFMSGBF()
                        Dim rawValue = lastClickedCell.Row.Cells("CUST_DEPT_CODE").Value
                        Dim deptValues = New String() {"Women", "Men"}
                        Dim deptCodes = New String() {"WM", "MN"}
                        Dim deptVal = input.Get_opt_from_User(String.Format("Select mapping for material ""{0}"":", rawValue), deptValues, 0, "Map Department")

                        If deptVal <> -1 Then
                            ASCDATA1.ExecuteSQL(
                            "MERGE INTO RSTIMAP1 RM " &
                            "USING (SELECT :PARM1 CUST_CODE, :PARM2 COLUMN_NAME, :PARM3 RAW_VALUE, :PARM4 IMPORT_VALUE FROM DUAL) VALS " &
                            "ON (RM.CUST_CODE=VALS.CUST_CODE AND RM.COLUMN_NAME=VALS.COLUMN_NAME AND RM.RAW_VALUE=VALS.RAW_VALUE) " &
                            "WHEN MATCHED THEN " &
                            "UPDATE SET RM.IMPORT_VALUE=VALS.IMPORT_VALUE " &
                            "WHEN NOT MATCHED THEN " &
                            "INSERT (CUST_CODE,COLUMN_NAME,RAW_VALUE,IMPORT_VALUE) VALUES (VALS.CUST_CODE,VALS.COLUMN_NAME,VALS.RAW_VALUE,VALS.IMPORT_VALUE)",
                            "VVVV", New Object() {custCode, "DEPT_CODE", rawValue, deptCodes(deptVal)})
                        End If

                    Case "MATL_CODE", "MATL_CODE_IMPORT"
                        Dim input As New ASFMSGBF()
                        Dim rawValue = lastClickedCell.Row.Cells("MATL_CODE").Value
                        Dim matlValues = New String() {"Silver", "Gold", "Cinta"}
                        Dim matlCodes = New String() {"S", "G", "C"}
                        Dim matlVal = input.Get_opt_from_User(String.Format("Select mapping for material ""{0}"":", rawValue), matlValues, 0, "Map Material")

                        If matlVal <> -1 Then
                            ASCDATA1.ExecuteSQL(
                            "MERGE INTO RSTIMAP1 RM " &
                            "USING (SELECT :PARM1 CUST_CODE, :PARM2 COLUMN_NAME, :PARM3 RAW_VALUE, :PARM4 IMPORT_VALUE FROM DUAL) VALS " &
                            "ON (RM.CUST_CODE=VALS.CUST_CODE AND RM.COLUMN_NAME=VALS.COLUMN_NAME AND RM.RAW_VALUE=VALS.RAW_VALUE) " &
                            "WHEN MATCHED THEN " &
                            "UPDATE SET RM.IMPORT_VALUE=VALS.IMPORT_VALUE " &
                            "WHEN NOT MATCHED THEN " &
                            "INSERT (CUST_CODE,COLUMN_NAME,RAW_VALUE,IMPORT_VALUE) VALUES (VALS.CUST_CODE,VALS.COLUMN_NAME,VALS.RAW_VALUE,VALS.IMPORT_VALUE)",
                            "VVVV", New Object() {custCode, "MATL_CODE", rawValue, matlCodes(matlVal)})
                        End If

                End Select
                dst.Tables("RSTIMPR2").Clear()
                Fill_Records("RSTIMPR2", New Object() {impNo})
        End Select
    End Sub
#End Region

    Private Function GetImportFileNames() As String()
        Dim getImportFile As New OpenFileDialog()
        getImportFile.Filter = "Excel/PDF Files|*.xls*;*.pdf"
        getImportFile.Multiselect = True
        If getImportFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Return getImportFile.FileNames
        End If
        Return Nothing
    End Function

    Private Sub optMode_ValueChanged(sender As System.Object, e As EventArgs) Handles optMode.ValueChanged
        Select Case optMode.Value
            Case Mode.RawFiles
                LoadImportFileHeaders(If(chkRawPreviouslyImported.Checked, ImportStatus.LoadedAlready, ImportStatus.NotLoaded))
            Case Mode.ImportedData
                LoadImportFileHeaders(optImportData.Value)
            Case Mode.IndependentEntry
                EntryMode = "L"
            Case Mode.NationalEntry
                EntryMode = "N"
            Case Mode.PlanEntry
                EntryMode = "P"
            Case Mode.ReportSettings
                EntryMode = "R"
                BindIndependentsFullList()
        End Select

        Mode_Settings(False)
    End Sub

    Private Sub chkRawPreviouslyImported_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkRawPreviouslyImported.CheckedChanged
        If chkRawPreviouslyImported.Checked Then
            grdRSTFILE1.Text = "Previously Imported"
            LoadRawFileHeaders()
        Else
            grdRSTFILE1.Text = "Raw Files"
            dst.Tables("RSTFILE1").Clear()
        End If
    End Sub

    Private Sub optImportData_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optImportData.ValueChanged
        LoadImportFileHeaders(optImportData.Value)
        Select Case optImportData.Value
            Case ImportStatus.NotLoaded
                grdRSTIMPR1.Text = "Not Loaded"
            Case ImportStatus.LoadedAlready
                grdRSTIMPR1.Text = "Loaded Already"
            Case ImportStatus.DeletedFromQueue
                grdRSTIMPR1.Text = "Deleted From Queue"
            Case ImportStatus.IndependentEntry
                grdRSTIMPR1.Text = "Manually Entered"
        End Select
        Mode_Settings(False)
    End Sub

    Private Sub btnSelectAll_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectAll.Click
        For Each grdRow In grdRSTIMPR1.Rows.GetFilteredInNonGroupByRows()
            'grdRow.Cells("SELECTED").Value = "1"
            dst.Tables("RSTIMPR1").Rows.Find(grdRow.Cells("IMPORT_NO").Value).Item("SELECTED") = "1"
        Next
    End Sub

    Private Sub btnDeSelectAll_Click(sender As System.Object, e As System.EventArgs) Handles btnDeSelectAll.Click
        For Each grdRow In grdRSTIMPR1.Rows
            'grdRow.Cells("SELECTED").Value = "0"
            dst.Tables("RSTIMPR1").Rows.Find(grdRow.Cells("IMPORT_NO").Value).Item("SELECTED") = "0"
        Next
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "REGION_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load")
                End If
        End Select
    End Sub

#Region "Void Matrices"
    Sub InitializeImportCountGrid(ByVal grd As UltraGrid)
        Dim W As Integer = 0
        Dim P As Integer = 0
        With grd.DisplayLayout.Bands(0)
            If .Groups.Count <> 0 Then
                For g As Integer = .Groups.Count - 1 To 0 Step -1
                    .Groups.Remove(g)
                Next
            End If
            .Groups.Add("CUST_CODE")
            .Groups("CUST_CODE").Header.Caption = ""
            .Columns("CUST_CODE").Group = .Groups("CUST_CODE")

            ASCMAIN1.sql = "SELECT * FROM GLTPARM3, " & _
                                "(SELECT TAPWKRA1(MAX(YYYYWW),-159) FIRST_WEEK,MAX(YYYYWW) LAST_WEEK FROM GLTPARM3 WHERE YYYYPP = :PARM1) WK " & _
                                " WHERE YYYYWW BETWEEN WK.FIRST_WEEK AND WK.LAST_WEEK"

            Dim wkTable As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ASCMAIN1.CYP})
            Dim periods As List(Of String) = GetDistinct("YYYYPP", wkTable)
            periods.Sort()
            For Each YYYYPP As String In periods
                P += 1
                .Groups.Add(YYYYPP)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YYYYPP)
                .Groups(YYYYPP).Header.Caption = Mid(LEGEND, 10, 6)
                .Groups(YYYYPP).Header.Appearance.BackColor = Color.Yellow
                .Groups(YYYYPP).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                For Each row As DataRow In wkTable.Select(String.Format("YYYYPP='{0}'", YYYYPP), "YYYYWW")
                    W += 1
                    Dim COLUMN_NAME As String = "YYYYWW_" & Format(W, "000")
                    .Columns.Add(COLUMN_NAME)
                    .Columns(COLUMN_NAME).DefaultCellValue = 0
                    .Columns(COLUMN_NAME).Format = "#0"
                    .Columns(COLUMN_NAME).Group = .Groups(YYYYPP)
                    Dim YW As String = row.Item("YYYYWW")
                    .Columns(COLUMN_NAME).Tag = YW
                    .Columns(COLUMN_NAME).Hidden = False
                    .Columns(COLUMN_NAME).Header.Caption = Mid(YW, 5, 2)
                    .Columns(COLUMN_NAME).Width = 30

                    If P Mod 2 = 0 Then
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightBlue
                    Else
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightPink
                    End If

                    Create_Summary(grd, COLUMN_NAME, "Sum")
                Next
            Next
        End With

        With grd.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
        End With

        Fill_Records(grd.DisplayLayout.Bands(0).Key, "", True)

        Sort_grdColumns(grd, "CUST_CODE")
    End Sub

    Sub RefreshImportCountGrid()
        If dst.Tables.Contains("RSTIMPCT") Then
            Select Case optImportCountType.Value
                Case "A"
                    grdRSTIMPCT.DataSource = dst.Tables("RSTIMPCT")
                    Fill_Records("RSTIMPCT")
                Case "W"
                    grdRSTIMPCT.DataSource = dst.Tables("RSTIMPCW")
                    Fill_Records("RSTIMPCW")
                Case "M"
                    grdRSTIMPCT.DataSource = dst.Tables("RSTIMPCM")
                    Fill_Records("RSTIMPCM")
            End Select
        End If
    End Sub

    Sub InitializeIndependentCountGrid(ByVal grd As UltraGrid)
        Dim P As Integer = 0
        With grd.DisplayLayout.Bands(0)
            If .Groups.Count <> 0 Then
                For g As Integer = .Groups.Count - 1 To 0 Step -1
                    .Groups.Remove(g)
                Next
            End If
            .Groups.Add("CUST_CODE")
            .Groups("CUST_CODE").Header.Caption = ""
            .Columns("CUST_CODE").Group = .Groups("CUST_CODE")
            .Columns("REGION_DESC").Group = .Groups("CUST_CODE")


            .Groups.Add("MONTHS")
            .Groups("MONTHS").Header.Caption = ""

            ASCMAIN1.sql = "SELECT * FROM GLTPARM2, " & _
                                "(SELECT TAPPRDA1(MAX(OPS_YYYYPP),-35) FIRST_MONTH,MAX(OPS_YYYYPP) LAST_MONTH FROM GLTPARM2 WHERE OPS_YYYYPP = :PARM1) MN " & _
                                " WHERE OPS_YYYYPP BETWEEN MN.FIRST_MONTH AND MN.LAST_MONTH"

            Dim perTable As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ASCMAIN1.CYP})
            Dim periods As List(Of String) = GetDistinct("OPS_YYYYPP", perTable)
            periods.Sort()
            For Each YYYYPP As String In periods
                P += 1
                '.Groups.Add(YYYYPP)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YYYYPP)
                '.Groups(YYYYPP).Header.Caption = Mid(LEGEND, 10, 6)
                '.Groups(YYYYPP).Header.Appearance.BackColor = Color.Yellow
                '.Groups(YYYYPP).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                Dim COLUMN_NAME As String = "YYYYPP_" & Format(P, "00")
                .Columns.Add(COLUMN_NAME)
                .Columns(COLUMN_NAME).DefaultCellValue = 0
                .Columns(COLUMN_NAME).Format = "#0"
                .Columns(COLUMN_NAME).Tag = YYYYPP
                .Columns(COLUMN_NAME).Hidden = If(P <= 12, True, False)
                .Columns(COLUMN_NAME).Header.Caption = Mid(LEGEND, 10, 6)
                .Columns(COLUMN_NAME).Width = 80
                .Columns(COLUMN_NAME).Group = .Groups("MONTHS")

                If P Mod 2 = 0 Then
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightBlue
                Else
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightPink
                End If

                Create_Summary(grd, COLUMN_NAME, "Sum")
            Next
        End With

        With grd.DisplayLayout.Bands(0)
            .Groups(0).Header.Fixed = True
        End With

        Fill_Records(grd.DisplayLayout.Bands(0).Key, "", True)

        Sort_grdColumns(grd, "CUST_CODE")
    End Sub

    Private Function GetDistinct(ByVal columnName As String, ByVal sourceTable As DataTable) As List(Of String)
        Dim view As New DataView(sourceTable)
        Dim departments As DataTable = view.ToTable(True, columnName)
        Return departments.AsEnumerable().Select(Function(row) row(columnName).ToString()).ToList()
    End Function

    Private Sub grdRSTIMPCT_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTIMPCT.InitializeLayout
        Dim redForeColor As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        redForeColor.ForeColor = Color.Red

        Dim equalsZero As Infragistics.Win.OperatorCondition = New OperatorCondition(ConditionOperator.Equals, 0)

        Dim conditionValueAppearance As New ConditionValueAppearance
        conditionValueAppearance.Add(equalsZero, redForeColor)

        For i As Integer = 1 To e.Layout.Bands(0).Columns.Count - 1
            e.Layout.Bands(0).Columns(i).ValueBasedAppearance = conditionValueAppearance
        Next
    End Sub

    Private Sub grdRSTIMPCT_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdRSTIMPCT.DoubleClickCell
        Try
            optMode.Value = Mode.ImportedData
            optImportData.Value = "0"
            Dim OPS_YYYYWW As String = e.Cell.Column.Tag
            Dim CUST_CODE As String = e.Cell.Row.Cells("CUST_CODE").Text
            With grdRSTIMPR1.DisplayLayout.Bands(0)
                .ColumnFilters.ClearAllFilters()
                .ColumnFilters("CUST_CODE").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.StartsWith, CUST_CODE.Substring(0, 4))
                .ColumnFilters("OPS_YYYYWW").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, OPS_YYYYWW)

                Show_Filter(grdRSTIMPR1, True)
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdRSTLOADX_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdRSTLOADX.DoubleClickCell
        Try
            optImportData.Value = "1"
            Dim OPS_YYYYWW As String = e.Cell.Column.Tag
            Dim CUST_CODE As String = e.Cell.Row.Cells("CUST_CODE").Text
            With grdRSTIMPR1.DisplayLayout.Bands(0)
                .ColumnFilters.ClearAllFilters()
                .ColumnFilters("CUST_CODE").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.StartsWith, CUST_CODE.Substring(0, 4))
                .ColumnFilters("OPS_YYYYWW").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, OPS_YYYYWW)

                Show_Filter(grdRSTIMPR1, True)
            End With
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "RSTIMPR1"

    Private Sub grdRSTIMPR1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdRSTIMPR1.BeforeRowsDeleted
        importNumbersForDelete.Clear()
        For Each grow As UltraGridRow In grdRSTIMPR1.Selected.Rows
            importNumbersForDelete.Add(grow.Cells("IMPORT_NO").Text)
        Next
    End Sub

    Private Sub grdRSTIMPR1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTIMPR1.AfterRowsDeleted
        Dim sql As String = ""
        For Each importNumber As String In importNumbersForDelete
            sql &= ",'" & importNumber & "'"
        Next
        ASCMAIN1.sql = "Update RSTIMPR1 Set IMPORT_STATUS = 'D' where IMPORT_NO in (" & Mid(sql, 2) & ")"
        ASCDATA1.ExecuteSQL()
        dst.Tables("RSTIMPR1").AcceptChanges()
        RefreshImportCountGrid()
        LoadImportFileHeaders(ImportStatus.NotLoaded)
    End Sub

    Private Sub grdRSTIMPR1_AfterSelectChange(sender As System.Object, e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdRSTIMPR1.AfterSelectChange
        If grdRSTIMPR1.Selected.Rows.Count = 1 Then
            Dim importNo As String = grdRSTIMPR1.Selected.Rows(0).Cells("IMPORT_NO").Value
            LoadImportFileDetails(importNo)
            grdRSTIMPR2.Text = String.Format("Raw File Data for Import {0}", importNo)
        End If
    End Sub

    Private Sub grdRSTIMPR1_DoubleClickRow(sender As System.Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTIMPR1.DoubleClickRow
        tabImportInfo.SelectedTab = tabImportInfo.Tabs("Raw Data")
    End Sub

#End Region

#Region "RSTIMPR2"
    Private Sub grdRSTIMPR2_InitializeRow(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTIMPR2.InitializeRow

        If e.Row.Cells("DEPT_CODE_IMPORT").Value & "" = "" Or e.Row.Cells("MATL_CODE_IMPORT").Value & "" = "" Then
            e.Row.Appearance.BackColor = Color.LightPink
        End If

        If e.Row.Cells("CUST_STORE_NO_IMPORT").Value & "" = "" And Not Regex.IsMatch(e.Row.Cells("CUST_STORE_NO").Value, "(\d{1,6}|DIRECT)") Then
            e.Row.Appearance.BackColor = Color.LightPink
        End If

    End Sub
#End Region

#Region "RSTSREG1"
    Private Sub grdRSTSREG1_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTSREG1.AfterRowInsert
        e.Row.Cells("CUST_CODE").Value = txtCustomersSM.Value
    End Sub
#End Region

#Region "RSTCUSTS"
    Private Sub grdRSTCUSTS_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTCUSTS.InitializeLayout
        e.Layout.Bands(0).Columns("CUST_REGION_CODE").ValueList = cmbRegionsSM
    End Sub


    Private Sub BindStoreMaintenanceDropDown(ByVal customerCode As String)
        Dim dt As DataTable = dst.Tables("RSTSREG1") ''ASCDATA1.GetDataTable("SELECT CUST_REGION_CODE,CUST_REGION_DESC FROM RSTSREG1 WHERE CUST_CODE=:PARM1", "", "V", New Object() {customerCode})

        cmbRegionsSM.SetDataBinding(dt, Nothing)
        cmbRegionsSM.ValueMember = "CUST_REGION_CODE"
        cmbRegionsSM.DisplayMember = "CUST_REGION_DESC"

        cmbRegionsSM.DisplayLayout.Bands(0).Columns("CUST_REGION_CODE").Hidden = True
        cmbRegionsSM.DisplayLayout.Bands(0).Columns("CUST_REGION_SEQ").Hidden = True
        cmbRegionsSM.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbRegionsSM.DisplayLayout.Bands(0).Columns("CUST_REGION_DESC").Header.Caption = "Region"

    End Sub

    Private Sub grdRSTCUSTS_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTCUSTS.AfterRowInsert
        e.Row.Cells("CUST_CODE").Value = txtCustomersSM.Value
    End Sub

    Private Sub grdRSTCUSTS_BeforeCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdRSTCUSTS.BeforeCellUpdate
        If e.Cell.Column.Key = "CUST_STORE_NO" Then
            If Not Regex.IsMatch(e.Cell.Text & "", "^\d+$") And e.Cell.Text.ToUpper() <> "DIRECT" Then 'only valid for nationals
                e.Cancel = True
                MsgBox("Invalid store #")
            End If
        End If
    End Sub

    Private Sub grdRSTCUSTS_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTCUSTS.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_STORE_NO" Then
                Dim newVal As String = e.Cell.Value.ToString().PadLeft(6, "0").ToUpper()
                e.Cell.SetValue(newVal, False)

                Dim storeName As String = ASCDATA1.GetDataValue("SELECT CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE=:PARM1 AND CUST_STORE_NO=:PARM2", "VV", New Object() {txtCustomersSM.Value, newVal}) & ""
                e.Cell.Row.Cells("CUST_STORE_NAME").SetValue(storeName, False)
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub
#End Region

#Region "RSTRETLI"



    Private Sub BindIndependentsDropDowns(ByVal regionCode As String)

        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT CUST_CODE,CUST_NAME FROM ARTCUST1 LEFT JOIN SOTSREP1 ON (ARTCUST1.SREP_CODE=SOTSREP1.SREP_CODE) WHERE ((NVL(TRADE_CLASS_CODE,'IND')='IND' AND :PARM1 NOT IN ('NEI','SAK')) OR ARTCUST1.CUST_CODE IN ('NMLASTCALL','SAKSOFF5TH')) AND NVL(SOTSREP1.SREP_CODE,'NA')=:PARM1", "", "V", New Object() {regionCode})

        cmbIndependentCustomers.SetDataBinding(dt, Nothing)
        cmbIndependentCustomers.ValueMember = "CUST_CODE"
        cmbIndependentCustomers.DisplayMember = "CUST_NAME"

        cmbIndependentCustomers.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbIndependentCustomers.DisplayLayout.Bands(0).Columns("CUST_NAME").Header.Caption = "Customer"
        cmbIndependentCustomers.DisplayLayout.PerformAutoResizeColumns(True, PerformAutoSizeType.AllRowsInBand)


        Dim dt2 As DataTable = ASCDATA1.GetDataTable("SELECT ARTCUST2.CUST_CODE,ARTCUST2.CUST_STORE_NO,ARTCUST2.CUST_STORE_NAME FROM ARTCUST1 JOIN ARTCUST2 ON (ARTCUST1.CUST_CODE=ARTCUST2.CUST_CODE) LEFT JOIN SOTSREP1 ON (ARTCUST1.SREP_CODE=SOTSREP1.SREP_CODE) WHERE (NVL(ARTCUST1.TRADE_CLASS_CODE,'IND')='IND' AND :PARM1 NOT IN ('SAK','NEI')) AND NVL(SOTSREP1.SREP_CODE,'NA')=:PARM1", "", "V", New Object() {regionCode})

        Select Case regionCode
            Case "NEI"
                dt2.Rows.Add(New Object() {"NMLASTCALL", "NMLASTCALL", "NM Last Call"})
            Case "SAK"
                dt2.Rows.Add(New Object() {"SAKSOFF5TH", "SAKSOFF5TH", "Saks Off 5th"})
        End Select

        cmbIndependentStores.SetDataBinding(dt2, Nothing)
        cmbIndependentStores.ValueMember = "CUST_STORE_NO"
        cmbIndependentStores.DisplayMember = "CUST_STORE_NAME"

        cmbIndependentStores.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbIndependentStores.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Hidden = True
        cmbIndependentStores.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Header.Caption = "Store"
        cmbIndependentStores.DisplayLayout.PerformAutoResizeColumns(True, PerformAutoSizeType.AllRowsInBand)
    End Sub

    Private Sub BindIndependentsFullList()
        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT CUST_CODE,CUST_NAME FROM ARTCUST1 WHERE NVL(TRADE_CLASS_CODE,'IND')='IND'")

        cmbAllIndependents.SetDataBinding(dt, Nothing)
        cmbAllIndependents.ValueMember = "CUST_CODE"
        cmbAllIndependents.DisplayMember = "CUST_NAME"

        cmbAllIndependents.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbAllIndependents.DisplayLayout.Bands(0).Columns("CUST_NAME").Header.Caption = "Customer"
        cmbAllIndependents.DisplayLayout.PerformAutoResizeColumns(True, PerformAutoSizeType.AllRowsInBand)
    End Sub

    Private Function CreateIndependentStoresValueList(ByVal custCode As String) As UltraCombo
        Dim cmb As New UltraCombo()
        Dim dt As New DataTable()
        dt = CType(cmbIndependentStores.DataSource, DataTable).Copy()
        Dim dv As DataView = dt.AsDataView()
        dv.RowFilter = String.Format("CUST_CODE='{0}'", custCode)

        cmb.SetDataBinding(dv, Nothing)
        cmb.ValueMember = "CUST_STORE_NO"
        cmb.DisplayMember = "CUST_STORE_NAME"

        cmb.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmb.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Hidden = True
        cmb.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Header.Caption = "Store"
        cmb.DisplayLayout.PerformAutoResizeColumns(True, PerformAutoSizeType.AllRowsInBand)
        cmb.DataSource = dv
        Return cmb
    End Function



    Private Sub grdRSTRETLI_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTRETLI.InitializeLayout
        e.Layout.Bands(0).Columns("CUST_CODE").ValueList = Me.cmbIndependentCustomers
    End Sub

    Private Sub grdRSTINDYS_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTINDYS.InitializeLayout
        e.Layout.Bands(0).Columns("CUST_CODE").ValueList = Me.cmbAllIndependents
    End Sub

    Private Sub grdRSTRETLI_BeforeCellListDropDown(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTRETLI.BeforeCellListDropDown
        If Not (e.Cell.Row.IsAddRow Or e.Cell.Row.IsFilterRow) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdRSTINDYS_BeforeCellListDropDown(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTINDYS.BeforeCellListDropDown
        If Not (e.Cell.Row.IsAddRow Or e.Cell.Row.IsFilterRow) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdRSTRETLI_BeforeRowActivate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETLI.BeforeRowActivate, grdRSTINDYS.BeforeRowActivate
        If e.Row.IsAddRow Or e.Row.IsFilterRow Then
            For Each cell In e.Row.Cells
                cell.IgnoreRowColActivation = True
                cell.Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            Next
        Else

        End If
    End Sub

    Private Sub grdRSTRETLI_InitializeRow(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTRETLI.InitializeRow
        'e.Row.Cells("CUST_STORE_NO").ValueList = cmbIndependentStores.data

        If e.Row.IsAddRow Or e.Row.IsFilterRow Then

        Else
            e.Row.Cells("CUST_STORE_NO").ValueList = CreateIndependentStoresValueList(e.Row.Cells("CUST_CODE").Value)
        End If
    End Sub

    Private Sub grdRSTRETLI_BeforeCellActivate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTRETLI.BeforeCellActivate
        If e.Cell.Column.Key = "CUST_STORE_NO" And e.Cell.Row.Cells("CUST_CODE").Value IsNot Nothing Then
            e.Cell.ValueList = CreateIndependentStoresValueList(e.Cell.Row.Cells("CUST_CODE").Value)
        End If
    End Sub

    Private Sub grdRSTRETLI_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLI.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        'Turn off the event to prevent recursion  
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_CODE" Then
                e.Cell.Row.Cells("CUST_STORE_NO").Value = DBNull.Value
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub

    Private Sub grdRSTRETLI_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETLI.AfterRowInsert
        e.Row.Cells("OPS_YYYYPP").Value = Absx1.cmbFor("RSRYP").Value
        e.Row.Cells("REGION_CODE").Value = cmbIndRegion.Value
        e.Row.Cells("DEPT_CODE").Value = "WM"
        e.Row.Cells("MATL_CODE").Value = "S"
    End Sub
#End Region

#Region "RSTPLAN1"

    Private Sub grdRSTPLAN1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTPLAN1.InitializeLayout
        Dim grd = CType(sender, UltraGrid)
        Dim departmentValueList As ValueList = grd.DisplayLayout.ValueLists.Add("VL1")

        departmentValueList.ValueListItems.Add("WM", "Women")
        departmentValueList.ValueListItems.Add("MN", "Men")

        Dim departmentColumn As UltraGridColumn = grd.DisplayLayout.Bands(0).Columns("DEPT_CODE")
        departmentColumn.ValueList = departmentValueList

        departmentColumn.Style = ColumnStyle.DropDownValidate
        departmentColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend

        Dim materialValueList As ValueList = grd.DisplayLayout.ValueLists.Add("VL2")

        materialValueList.ValueListItems.Add("S", "Silver")
        materialValueList.ValueListItems.Add("G", "Gold")
        materialValueList.ValueListItems.Add("C", "Cinta")

        Dim materialColumn As UltraGridColumn = grd.DisplayLayout.Bands(0).Columns("MATL_CODE")
        materialColumn.ValueList = materialValueList

        materialColumn.Style = ColumnStyle.DropDownValidate
        materialColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend

        Dim redForeColor As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        redForeColor.BackColor = Color.Red

        Dim equalsZero As Infragistics.Win.OperatorCondition = New OperatorCondition(ConditionOperator.IsNullOrEmpty, "")

        Dim conditionValueAppearance As New ConditionValueAppearance
        conditionValueAppearance.Add(equalsZero, redForeColor)

        For i As Integer = 1 To e.Layout.Bands(0).Columns.Count - 1
            If e.Layout.Bands(0).Columns(i).Key = "MATL_CODE" Or e.Layout.Bands(0).Columns(i).Key = "DEPT_CODE" Then
                e.Layout.Bands(0).Columns(i).ValueBasedAppearance = conditionValueAppearance
            End If
        Next
    End Sub

    Private Sub grdRSTPLAN1_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTPLAN1.AfterRowInsert
        e.Row.Cells("CUST_CODE").Value = cmbPlanCustomers.Value
        e.Row.Cells("OPS_YYYYPP").Value = Absx1.txtFor("OPS_YYYYPP").Text
    End Sub

    Private Sub grdRSTPLAN1_BeforeCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdRSTPLAN1.BeforeCellUpdate
        If e.Cell.Column.Key = "CUST_STORE_NO" Then
            If e.Cell.Text <> "" And Not Regex.IsMatch(e.Cell.Text & "", "^\d+$") And e.Cell.Text.ToUpper() <> "DIRECT" Then
                e.Cancel = True
                MsgBox("Invalid store #")
            End If
        End If
    End Sub

    Private Sub grdRSTPLAN1_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTPLAN1.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_STORE_NO" Then
                Dim newVal As String = e.Cell.Value.ToString().PadLeft(6, "0").ToUpper()
                e.Cell.SetValue(newVal, False)

                Dim storeName As String = ASCDATA1.GetDataValue("SELECT CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE=:PARM1 AND CUST_STORE_NO=:PARM2", "VV", New Object() {cmbPlanCustomers.Value, newVal}) & ""
                e.Cell.Row.Cells("CUST_STORE_NAME").SetValue(storeName, False)
            End If

            If e.Cell.Column.Key = "MATL_CODE" Then
                Select Case e.Cell.Value.ToString.ToUpper()
                    Case "SILVER", "S"
                        e.Cell.SetValue("S", False)
                    Case "GOLD", "G"
                        e.Cell.SetValue("G", False)
                    Case "CINTA", "C"
                        e.Cell.SetValue("C", False)
                    Case Else
                        e.Cell.Value = ""
                End Select
            End If

            If e.Cell.Column.Key = "DEPT_CODE" Then
                Select Case e.Cell.Value.ToString.ToUpper()
                    Case "MEN", "MN"
                        e.Cell.SetValue("MN", False)
                    Case "WOMEN", "WM"
                        e.Cell.SetValue("WM", False)
                    Case Else
                        e.Cell.Value = ""
                End Select
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub
#End Region

#Region "RSTPLAN6"
    Private Sub grdRSTPLAN6_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTPLAN6.AfterRowInsert
        e.Row.Cells("CUST_CODE").Value = cmbPlanCustomers.Value
    End Sub

    Private Sub grdRSTPLAN6_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTPLAN6.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_STORE_NO" Then
                Dim newVal As String = e.Cell.Value.ToString().PadLeft(6, "0").ToUpper()
                e.Cell.SetValue(newVal, False)
            End If

            If e.Cell.Column.Key = "MATL_CODE" Then
                Select Case e.Cell.Value.ToString.ToUpper()
                    Case "SILVER", "S"
                        e.Cell.SetValue("S", False)
                    Case "GOLD", "G"
                        e.Cell.SetValue("G", False)
                    Case "CINTA", "C"
                        e.Cell.SetValue("C", False)
                    Case Else
                        e.Cell.Value = ""
                End Select
            End If

            If e.Cell.Column.Key = "DEPT_CODE" Then
                Select Case e.Cell.Value.ToString.ToUpper()
                    Case "MEN", "MN"
                        e.Cell.SetValue("MN", False)
                    Case "WOMEN", "WM"
                        e.Cell.SetValue("WM", False)
                    Case Else
                        e.Cell.Value = ""
                End Select
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub

    Private Sub grdRSTPLAN6_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTPLAN6.InitializeLayout
        Dim departmentValueList As ValueList = Me.grdRSTPLAN6.DisplayLayout.ValueLists.Add("VL1")
        departmentValueList.ValueListItems.Add("WM", "Women")
        departmentValueList.ValueListItems.Add("MN", "Men")

        Dim departmentColumn As UltraGridColumn = Me.grdRSTPLAN6.DisplayLayout.Bands(0).Columns("DEPT_CODE")
        departmentColumn.ValueList = departmentValueList

        departmentColumn.Style = ColumnStyle.DropDownValidate
        departmentColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend


        Dim materialValueList As ValueList = grdRSTPLAN6.DisplayLayout.ValueLists.Add("VL2")

        materialValueList.ValueListItems.Add("S", "Silver")
        materialValueList.ValueListItems.Add("G", "Gold")
        materialValueList.ValueListItems.Add("C", "Cinta")

        Dim materialColumn As UltraGridColumn = Me.grdRSTPLAN6.DisplayLayout.Bands(0).Columns("MATL_CODE")
        materialColumn.ValueList = materialValueList

        materialColumn.Style = ColumnStyle.DropDownValidate
        materialColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend
    End Sub
#End Region

    Private Sub BindJHPlanCustomerNames()
        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT CUST_CODE,CASE CUST_CODE WHEN 'NMLASTCALL' THEN 'Neiman Marcus Last Call' WHEN 'SAKSOFF5TH' THEN 'Saks Off 5th' ELSE CUST_NAME END CUST_NAME FROM ARTCUST1 WHERE CUST_CODE IN (SELECT CUST_CODE FROM RSTINDYS) OR CUST_CODE IN ('NEIMANM10','HOLTREN10','BLOOMIES10','NORDSTR10','SAKSFIF10','NMLASTCALL','SAKSOFF5TH','ECOMSALE10')")
        cmbJHPlanCustomers.SetDataBinding(dt, Nothing)
        cmbJHPlanCustomers.ValueMember = "CUST_CODE"
        cmbJHPlanCustomers.DisplayMember = "CUST_NAME"

        cmbJHPlanCustomers.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        cmbJHPlanCustomers.DisplayLayout.Bands(0).Columns("CUST_NAME").Header.Caption = "Customer Name"
        cmbJHPlanCustomers.DisplayLayout.Bands(0).Columns("CUST_NAME").Width = 220
    End Sub

#Region "RSTRETL5"

    Private Sub BindNationalStoreNames(ByVal customerCode As String)
        Dim dt As DataTable = ASCDATA1.GetDataTable("SELECT CUST_STORE_NAME,CUST_STORE_NAME || ' (' || LPAD(CUST_STORE_NO,6,'0') || ')' CUST_STORE_DISP FROM ARTCUST2 WHERE CUST_CODE=:PARM1 AND CUST_STORE_NAME IS NOT NULL", "", "V", New Object() {customerCode})
        cmbNationalStores.SetDataBinding(dt, Nothing)
        cmbNationalStores.ValueMember = "CUST_STORE_NAME"
        cmbNationalStores.DisplayMember = "CUST_STORE_DISP"

        cmbNationalStores.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME").Hidden = True
        cmbNationalStores.DisplayLayout.Bands(0).Columns("CUST_STORE_DISP").Header.Caption = "Store Name"
        If cmbNationalStores.DisplayLayout.Bands(0).Columns("CUST_STORE_DISP").Width < 150 Then
            cmbNationalStores.DisplayLayout.Bands(0).Columns("CUST_STORE_DISP").Width *= 4
        End If
    End Sub

    Private Sub grdRSTRETL5_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTRETL5.InitializeLayout
        Dim departmentValueList As ValueList = Me.grdRSTRETL5.DisplayLayout.ValueLists.Add("VL1")

        departmentValueList.ValueListItems.Add("WM", "Women")
        departmentValueList.ValueListItems.Add("MN", "Men")

        Dim departmentColumn As UltraGridColumn = Me.grdRSTRETL5.DisplayLayout.Bands(0).Columns("DEPT_CODE")
        departmentColumn.ValueList = departmentValueList

        departmentColumn.Style = ColumnStyle.DropDownValidate
        departmentColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend

        Dim materialValueList As ValueList = Me.grdRSTRETL5.DisplayLayout.ValueLists.Add("VL2")

        materialValueList.ValueListItems.Add("S", "Silver")
        materialValueList.ValueListItems.Add("G", "Gold")
        materialValueList.ValueListItems.Add("C", "Cinta")

        Dim materialColumn As UltraGridColumn = Me.grdRSTRETL5.DisplayLayout.Bands(0).Columns("MATL_CODE")
        materialColumn.ValueList = materialValueList
        materialColumn.Style = ColumnStyle.DropDownValidate
        materialColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend

        Dim storeNamesColumn As UltraGridColumn = Me.grdRSTRETL5.DisplayLayout.Bands(0).Columns("CUST_STORE_NAME")
        storeNamesColumn.ValueList = Me.cmbNationalStores
        storeNamesColumn.Style = ColumnStyle.DropDownValidate
        storeNamesColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend
    End Sub

    Private Sub grdRSTRETL5_BeforeCellListDropDown(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdRSTRETL5.BeforeCellListDropDown
        If Not (e.Cell.Row.IsAddRow Or e.Cell.Row.IsFilterRow) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdRSTRETL5_BeforeRowActivate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETL5.BeforeRowActivate
        If e.Row.IsAddRow Or e.Row.IsFilterRow Then
            For Each cell In e.Row.Cells
                cell.IgnoreRowColActivation = True
                cell.Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            Next
        End If
    End Sub

    Private Sub grdRSTRETL5_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETL5.AfterRowInsert
        e.Row.Cells("IMPORT_NO").Value = editImportNumber
        e.Row.Cells("CUST_CODE").Value = cmbCustomersNationalEntry.Value
        If cmbCustomersNationalEntry.Value = "ECOMSALE10" Then
            e.Row.Cells("CUST_STORE_NO").Value = "ECOMSALE10"
        End If
        e.Row.Cells("OPS_YYYYWW").Value = GetYYYYWWFromDropDown(Absx1.cmbFor("RSRYW").Value, cmbCustomersNationalEntry.Value)
        e.Row.Cells("OPS_YYYYPP").Value = ASCDATA1.GetDataValue("SELECT YYYYPP FROM GLTPARM3 WHERE YYYYWW=:PARM1", "V", New Object() {GetYYYYWWFromDropDown(Absx1.cmbFor("RSRYW").Value, cmbCustomersNationalEntry.Value)})

        With grdRSTRETL5.DisplayLayout.Bands(0)
            .ColumnFilters.ClearAllFilters()
            Show_Filter(grdRSTIMPR1, False)
        End With
    End Sub

    Private Sub grdRSTRETL5_BeforeCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdRSTRETL5.BeforeCellUpdate
        If e.Cell.Column.Key = "CUST_STORE_NO" Then

            If Not Regex.IsMatch(e.Cell.Text & "", "^\d+$") And e.Cell.Text.ToUpper() <> "DIRECT" And cmbCustomersNationalEntry.Value <> "ECOMSALE10" Then
                e.Cancel = True
                MsgBox("Invalid store #")
            End If
        End If
    End Sub

    Private Sub grdRSTRETL5_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETL5.AfterCellUpdate
        Dim grid As UltraGrid = CType(sender, UltraGrid)
        grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, False)

        Try
            If e.Cell.Column.Key = "CUST_STORE_NO" Then
                Dim newVal As String = e.Cell.Value.ToString().PadLeft(6, "0").ToUpper()
                e.Cell.SetValue(newVal, False)

                Dim storeName As String = ASCDATA1.GetDataValue("SELECT CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_CODE=:PARM1 AND CUST_STORE_NO=:PARM2", "VV", New Object() {cmbCustomersNationalEntry.Value, newVal}) & ""
                e.Cell.Row.Cells("CUST_STORE_NAME").SetValue(storeName, False)

                With grdRSTRETL5.DisplayLayout.Bands(0)
                    .ColumnFilters.ClearAllFilters()
                    .ColumnFilters("CUST_STORE_NO").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, newVal)

                    Show_Filter(grdRSTIMPR1, True)
                End With
            ElseIf e.Cell.Column.Key = "CUST_STORE_NAME" Then
                Dim parenIndex As Integer = e.Cell.Text.LastIndexOf("(")
                Dim storeNo As String = e.Cell.Text.Substring(parenIndex + 1, 6)
                grid.EventManager.SetEnabled(GridEventIds.BeforeCellUpdate, False)
                e.Cell.Row.Cells("CUST_STORE_NO").SetValue(storeNo, False)
                grid.EventManager.SetEnabled(GridEventIds.BeforeCellUpdate, True)
            End If
        Finally
            grid.EventManager.SetEnabled(GridEventIds.AfterCellUpdate, True)
        End Try
    End Sub


    Private Sub grdRSTRETL5_BeforeRowUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETL5.BeforeRowUpdate
        If chkCinta.Checked = True AndAlso e.Row.Cells("MATL_CODE").Value & "" = "C" Then 'Remove this amount from the corresponding Gold row
            Try
                Dim goldRow As DataRow = dst.Tables("RSTRETL5").Select(String.Format("CUST_STORE_NO='{0}' AND DEPT_CODE='{1}' AND MATL_CODE='G'", e.Row.Cells("CUST_STORE_NO").Value, e.Row.Cells("DEPT_CODE").Value))(0)
                goldRow.Item("AMT_SOLD") -= (e.Row.Cells("AMT_SOLD").Value - Val(e.Row.Cells("AMT_SOLD").OriginalValue & ""))
            Catch ex As Exception
                e.Row.CancelUpdate()
                MsgBox("No corresponding gold entry found")
            End Try
        End If
        grdRSTRETL5.DisplayLayout.RowScrollRegions(0).ScrollRowIntoView(e.Row)
    End Sub

#End Region

    Private Sub UltraOptionSet1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optImportCountType.ValueChanged
        RefreshImportCountGrid()
    End Sub

    Private Sub grdRSTIMPR2_InitializeLayout(ByVal sender As Object, ByVal e As InitializeLayoutEventArgs) Handles grdRSTIMPR2.InitializeLayout
        Dim departmentValueList As ValueList = grdRSTIMPR2.DisplayLayout.ValueLists.Add("VL1")

        departmentValueList.ValueListItems.Add("WM", "Women")
        departmentValueList.ValueListItems.Add("MN", "Men")

        Dim departmentColumn As UltraGridColumn = grdRSTIMPR2.DisplayLayout.Bands(0).Columns("DEPT_CODE_IMPORT")
        departmentColumn.ValueList = departmentValueList
        departmentColumn.CellActivation = Activation.NoEdit

        Dim materialValueList As ValueList = grdRSTIMPR2.DisplayLayout.ValueLists.Add("VL2")

        materialValueList.ValueListItems.Add("S", "Silver")
        materialValueList.ValueListItems.Add("G", "Gold")
        materialValueList.ValueListItems.Add("C", "Cinta")

        Dim materialColumn As UltraGridColumn = grdRSTIMPR2.DisplayLayout.Bands(0).Columns("MATL_CODE_IMPORT")
        materialColumn.ValueList = materialValueList
        materialColumn.CellActivation = Activation.NoEdit
    End Sub

    Private Sub grdRSTIMPR2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdRSTIMPR2.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim mousePoint As Point = New Point(e.X, e.Y)
            'get the user interface element from the location of the mouse click
            Dim element As UIElement = grdRSTIMPR2.DisplayLayout.UIElement.ElementFromPoint(mousePoint)
            lastClickedCell = GetCell(element)
        End If
    End Sub

    Private Function GetCell(ByVal element As UIElement) As UltraGridCell
        If (element Is Nothing Or element.Parent Is Nothing) Then
            Return Nothing
        End If

        If (TypeOf element.Parent Is CellUIElement) Then
            Return CType(element.Parent, CellUIElement).Cell
        Else
            Return GetCell(element.Parent)
        End If
    End Function

    Private Sub optPlanType_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPlanType.ValueChanged
        pnlPlanOptions.Visible = (optPlanType.Value = "N")
    End Sub

    Private Sub grdRSTPLANJ_AfterRowInsert(sender As System.Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTPLANJ.AfterRowInsert
        e.Row.Cells("YYYY").Value = cmbPlanYear.Value
        e.Row.Cells("DIRECT").Value = "0"
    End Sub

    Private Sub grdRSTPLANJ_AfterCellUpdate(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTPLANJ.AfterCellUpdate
        If e.Cell.Tag = "CUST_CODE" Then
            If 1 = 2 Then 'IF THIS IS AN INDY, SKIP TO AMT SOLD
            End If
        End If
    End Sub


    Private Sub grdRSTPLANJ_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdRSTPLANJ.InitializeLayout
        Dim departmentValueList As ValueList = Me.grdRSTPLANJ.DisplayLayout.ValueLists.Add("VL1")
        departmentValueList.ValueListItems.Add("WM", "Women")
        departmentValueList.ValueListItems.Add("MN", "Men")

        Dim departmentColumn As UltraGridColumn = Me.grdRSTPLANJ.DisplayLayout.Bands(0).Columns("DEPT_CODE")
        departmentColumn.ValueList = departmentValueList

        departmentColumn.Style = ColumnStyle.DropDownValidate
        departmentColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend


        Dim custCodeColumn As UltraGridColumn = Me.grdRSTPLANJ.DisplayLayout.Bands(0).Columns("CUST_CODE")
        custCodeColumn.ValueList = Me.cmbJHPlanCustomers
        custCodeColumn.Style = ColumnStyle.DropDownValidate
        custCodeColumn.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend
    End Sub

    Private Sub grdRSTCUSTS_ClickCellButton(sender As System.Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTCUSTS.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim SQL As String = ""

        If e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit Then
            Exit Sub
        End If

        Select Case COLUMN_NAME
            Case "YYYYPP_DOOR_OPENED", "YYYYPP_DOOR_CLOSED", "YYYYPP_DOOR_OPENED_M", "YYYYPP_DOOR_CLOSED_M"
                ' If you need to limit the select then add a where clause
                SQL = "SELECT LEGEND, OPS_YYYYPP FROM GLTPARM2"
        End Select

        ASCMAIN1.CodeSelector.SQL = SQL
        ASCMAIN1.CodeSelector.VIEW_NAME = ""
        ASCMAIN1.CodeSelector.MultipleSelections = False
        ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = e.Cell.Text
        Dim F As New ASFCODE1
        F.ShowDialog()
        F.Dispose()

        If ASCMAIN1.CodeSelector.Selections <> 0 Then

            Select Case COLUMN_NAME
                Case "YYYYPP_DOOR_OPENED", "YYYYPP_DOOR_CLOSED", "YYYYPP_DOOR_OPENED_M", "YYYYPP_DOOR_CLOSED_M"
                    grdRSTCUSTS.ActiveRow.Cells(COLUMN_NAME).Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item(1)
            End Select
        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                sql_where = "CUST_STATUS = 'A' AND TRADE_CLASS_CODE IN ('IND','NAT')"
        End Select
    End Sub

    Private Sub chkShowClosedDoors_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowClosedDoors.CheckedChanged
        If chkShowClosedDoors.Checked Then
            grdRSTRETLI.DisplayLayout.Bands(0).ColumnFilters("OPEN_DOOR").ClearFilterConditions()
        Else
            grdRSTRETLI.DisplayLayout.Bands(0).ColumnFilters("OPEN_DOOR").FilterConditions.Add(FilterComparisionOperator.Equals, "1")
        End If
    End Sub

    Private Sub tabImportMode_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabImportMode.SelectedTabChanged
        If tabImportMode.SelectedTab IsNot Nothing AndAlso tabImportMode.SelectedTab.Key = Mode.ReportSettings Then
            Fill_Record("RSTINDYS")
        End If
    End Sub

    Private Sub cmbAllIndependents_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles cmbAllIndependents.KeyPress
        If e.KeyChar = Chr(13) Then

        End If
    End Sub
End Class

Structure ImportStatus
    Const NotLoaded = "0"
    Const LoadedAlready = "1"
    Const DeletedFromQueue = "D"
    Const IndependentEntry = "M"
End Structure

Structure Mode
    Const RawFiles = "Raw Files"
    Const ImportedData = "Imported Data"
    Const NationalEntry = "National Entry"
    Const IndependentEntry = "Independent Entry"
    Const PlanEntry = "Plan Entry"
    Const StoreMaintenance = "Store Maintenance"
    Const ReportSettings = "Report Settings"
End Structure
