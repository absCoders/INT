Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports System.Net
Imports Newtonsoft.Json.Linq
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions
Imports System.IO
Public Class ICFSLPI1

    Dim FILENAME As String = ""

    Dim ABS2SLP As New Dictionary(Of String, String)
    Dim SLP2ABS As New Dictionary(Of String, String)

    Dim ABS_TABLES As New List(Of String)

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing
    Dim ARCHIVE_DIR As String = "Z:\SLP\Share\SLP\XLS\Item Master Import Archive"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ABS2SLP.Add("ITEM_STATUS", "Status")
        ABS2SLP.Add("ITEM_UPC_CODE", "UPC")
        ABS2SLP.Add("ITEM_CODE", "UPC")
        ABS2SLP.Add("STYLE_CODE", "SKU*")
        ABS2SLP.Add("ITEM_DESC", "Product Description")
        ABS2SLP.Add("HC_CODE", "Product Category")
        ABS2SLP.Add("COLLECTION_CODE", "Product Sub-Category 1")
        ABS2SLP.Add("ITEM_CLASS_CODE", "Product Sub-Category 2")
        ABS2SLP.Add("PROD_CODE", "Product Type")
        ABS2SLP.Add("SEASON_CODE", "Release")
        ABS2SLP.Add("ITEM_CATGY_CODE", "Core / Limited Edition")
        ABS2SLP.Add("ITEM_SHIP_DATE", "ATS Date")
        ABS2SLP.Add("ITEM_LAUNCH_DATE", "Wholesale Launch Date")
        ABS2SLP.Add("ITEM_STD_PACK_SLS", "Inner Carton Quantity")
        ABS2SLP.Add("CARTON_PACK_QTY", "Master Carton Quantity")
        ABS2SLP.Add("ITEM_WEIGHT", "Unit Gross Weight (LBS/pounds")
        ABS2SLP.Add("ITEM_UNIT_LENGTH", "Unit Length (IN/inches)")
        ABS2SLP.Add("ITEM_UNIT_WIDTH", "Unit Width (IN/inches)")
        ABS2SLP.Add("ITEM_UNIT_HEIGHT", "Unit Height (IN/inches)")
        ABS2SLP.Add("ITEM_PALLET_WEIGHT", "Inner Carton Gross Weight (LBS/pounds)")
        ABS2SLP.Add("ITEM_PALLET_LENGTH", "Inner Carton Length (IN/inches)")
        ABS2SLP.Add("ITEM_PALLET_WIDTH", "Inner Carton Width (IN/inches)")
        ABS2SLP.Add("ITEM_PALLET_HEIGHT", "Inner Carton Height (IN/inches_)")
        ABS2SLP.Add("CASE_WEIGHT_GRS", "Master Carton Gross Weight (LBS/pounds)")
        ABS2SLP.Add("ITEM_CASE_LENGTH", "Master Carton Length (IN/inches)")
        ABS2SLP.Add("ITEM_CASE_WIDTH", "Master Carton Width (IN/inches)")
        ABS2SLP.Add("ITEM_CASE_HEIGHT", "Master Carton Height (IN/inches)")
        ABS2SLP.Add("ITEM_RETAIL_PRICE", "RRP - US(USD $) - Current")
        ABS2SLP.Add("ITEM_PRICE", "RRP - CA(CAD $) - Current")

        'For Each key As String In ABS2SLP.Keys
        '    SLP2ABS.Add(ABS2SLP(key), key)
        'Next


        With dst

            ABS_TABLES.Add("ICTITEM1")
            ABS_TABLES.Add("ICTCOLL1")
            ABS_TABLES.Add("ICTPROD1")
            ABS_TABLES.Add("ICTSEAS1")
            ABS_TABLES.Add("ICTCLAS1")
            ABS_TABLES.Add("ICTCOLL0")
            'ABS_TABLES.Add("ARTCLAS1")
            'ABS_TABLES.Add("TATTERM1")

            Create_TDA(.Tables.Add, "ICTRETLC", "*", 0)
            Create_TDA(.Tables.Add, "ICTVALUC", "*", 0)

            For Each TABLE_NAME As String In ABS_TABLES
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0)
                AUDIT.Add(TABLE_NAME, "*")
            Next

            With .Tables.Add("ARTSLPE1")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("COLUMN_NAME")
                '.Columns.Add("ERROR_CODE")
                .Columns.Add("ERROR")

                '.PrimaryKey = New DataColumn() { .Columns("TABLE_NAME"), .Columns("COLUMN_NAME"), .Columns("ERROR_CODE")}
            End With

            With .Tables.Add("ICTITEMX")
                For Each COLUMN_NAME As String In ABS2SLP.Keys
                    .Columns.Add(COLUMN_NAME)
                Next
                .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE")}
            End With

            With .Tables.Add("ICTCOLLX")
                .Columns.Add("COLLECTION_CODE")
                .Columns.Add("COLLECTION_CODE_ABS")
                .PrimaryKey = New DataColumn() { .Columns("COLLECTION_CODE")}
            End With

            With .Tables.Add("ICTPRODX")
                .Columns.Add("PROD_CODE")
                .Columns.Add("PROD_CODE_ABS")
                .PrimaryKey = New DataColumn() { .Columns("PROD_CODE")}
            End With

            With .Tables.Add("ICTSEASX")
                .Columns.Add("SEASON_CODE")
                .Columns.Add("SEASON_CODE_ABS")
                .PrimaryKey = New DataColumn() { .Columns("SEASON_CODE")}
            End With

            With .Tables.Add("ICTCLASX")
                .Columns.Add("ITEM_CLASS_CODE")
                .Columns.Add("ITEM_CLASS_CODE_ABS")
                .PrimaryKey = New DataColumn() { .Columns("ITEM_CLASS_CODE")}
            End With

            Create_TDA(.Tables.Add, "SLTCOLL1", "*", 0, False)
            Fill_Records("SLTCOLL1")
            With .Tables("SLTCOLL1")
                .PrimaryKey = New DataColumn() { .Columns("COLLECTION_NAME")}
            End With

            If grdICTCOLLX.DisplayLayout.Bands.Count > 0 AndAlso grdICTCOLLX.DisplayLayout.Bands(0).Columns.Exists("COLLECTION_CODE_ABS") Then
                With grdICTCOLLX.DisplayLayout.Bands(0).Columns("COLLECTION_CODE_ABS")
                    .MaxLength = 6
                    .CharacterCasing = CharacterCasing.Upper
                End With
            End If
            If grdICTPRODX.DisplayLayout.Bands.Count > 0 AndAlso grdICTPRODX.DisplayLayout.Bands(0).Columns.Exists("PROD_CODE_ABS") Then
                With grdICTPRODX.DisplayLayout.Bands(0).Columns("PROD_CODE_ABS")
                    .MaxLength = 6
                    .CharacterCasing = CharacterCasing.Upper
                End With
            End If

        End With

        grdICTITEMX.DataSource = dst.Tables("ICTITEMX")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")
        grdICTCOLL1.DataSource = dst.Tables("ICTCOLL1")
        grdICTCOLLX.DataSource = dst.Tables("ICTCOLLX")
        grdICTPROD1.DataSource = dst.Tables("ICTPROD1")
        grdICTPRODX.DataSource = dst.Tables("ICTPRODX")
        grdICTSEAS1.DataSource = dst.Tables("ICTSEAS1")
        grdICTSEASX.DataSource = dst.Tables("ICTSEASX")
        grdICTCLAS1.DataSource = dst.Tables("ICTCLAS1")
        grdICTCLASX.DataSource = dst.Tables("ICTCLASX")
        grdARTSLPE1.DataSource = dst.Tables("ARTSLPE1")


        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTITEM1.DisplayLayout.Bands(0).Columns
            Dim C As String = gcol.Key
            If ABS2SLP.ContainsKey(C) Then
            Else
                gcol.Hidden = True
            End If
        Next

        dst.Tables("ICTCOLLX").Columns("COLLECTION_CODE_ABS").ReadOnly = False
        dst.Tables("ICTPRODX").Columns("PROD_CODE_ABS").ReadOnly = False
        With grdICTCOLLX.DisplayLayout
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            .Override.CellClickAction = UltraWinGrid.CellClickAction.Edit
            If .Bands.Count > 0 Then
                With .Bands(0)
                    .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                    If .Columns.Exists("COLLECTION_CODE_ABS") Then
                        .Columns("COLLECTION_CODE_ABS").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End With
            End If
        End With

        With grdICTPRODX.DisplayLayout
            .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            .Override.CellClickAction = UltraWinGrid.CellClickAction.Edit
            If .Bands.Count > 0 Then
                With .Bands(0)
                    .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                    If .Columns.Exists("PROD_CODE_ABS") Then
                        .Columns("PROD_CODE_ABS").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End With
            End If
        End With

        Create_Summary(grdICTITEMX, "ITEM_CODE", "Count")
        Create_Summary(grdICTITEM1, "ITEM_CODE", "Count")


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Locate the workbook containing the data to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsX files (*.xlsx)|*.xlsx"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                    Else
                        Exit Sub
                    End If
                End Using

            Case "Update"
                If dst.Tables("ARTSLPE1").Rows.Count > 0 Then
                    EMsg &= vbCr & "Cannot Update - resolve errors"
                    tabData.Tabs("Errors").Selected = True
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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)


                'Case "Print"
                '    Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabData.Visible = tf
        '   UltraExplorerBarContainerControl1.Enabled = Not tf

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTITEMX", "ICTITEM1", "ICTCOLLX", "ICTCOLL1", "ICTPRODX", "ICTPROD1", "ICTSEASX", "ICTSEAS1", "ICTCLAS1", "ICTCLASX", "ARTSLPE1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data from Database")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        For Each TABLE_NAME As String In ABS_TABLES
            Fill_Records(TABLE_NAME)
        Next

        WorkbookView1.GetLock()
        ' WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

        workbook = WorkbookView1.ActiveWorkbook
        worksheet = workbook.Worksheets(0)

        Dim rH As Integer = -1 ' 1 ' 0 ' 1
        If worksheet.Cells(0, 0).Value = "Status" Then
            rH = 0
        ElseIf worksheet.Cells(1, 0).Value = "Status" Then
            rH = 1
        Else
            rH = 0
            MsgBox("Expected Column for Status to appear in Row 1 or Row 2")
            dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTITEM1", "*", $"Expected Column for Status to appear in Row 1 or Row 2"})
        End If

        Dim SLPXLS2COL As New Dictionary(Of String, Int32)

        'For Each key As String In ABS2SLP.Keys
        '    SLP2ABS.Add(ABS2SLP(key), key)
        'Next
        For c As Int32 = 0 To worksheet.UsedRange.ColumnCount - 1
            Dim cH As String = (worksheet.Cells(rH, c).Value & "").Trim()
            If cH <> "" Then
                For Each absCol As String In ABS2SLP.Keys
                    If String.Equals(ABS2SLP(absCol), cH, StringComparison.InvariantCultureIgnoreCase) Then
                        If Not SLPXLS2COL.ContainsKey(absCol) Then
                            SLPXLS2COL.Add(absCol, c)
                        End If
                    End If
                Next
            End If
        Next

        dst.Tables("ICTITEMX").Rows.Clear()

        For r As Int32 = rH + 1 To worksheet.UsedRange.RowCount - 1
            Dim A As String = worksheet.Cells(r, 0).Value & "" ' Col A value - stop when empty
            Dim B As String = worksheet.Cells(r, 1).Value & "" ' Col B value - stop when empty
            If A = "" And B = "" Then Exit For ' Stop
            Dim rowICTITEMX As DataRow = dst.Tables("ICTITEMX").NewRow
            With rowICTITEMX
                For Each key As String In SLPXLS2COL.Keys
                    Dim XC As Int32 = SLPXLS2COL(key)
                    .Item(key) = worksheet.Cells(r, XC).Value & ""
                Next
                dst.Tables("ICTITEMX").Rows.Add(rowICTITEMX)
            End With
        Next
        Sort_grdColumns(grdICTITEMX, "ITEM_CODE")
        Dim ITEM_CODE_col As Integer = If(SLPXLS2COL.ContainsKey("ITEM_CODE"), SLPXLS2COL("ITEM_CODE"), 0)
        Write_ICTRETLC(worksheet, rH, ITEM_CODE_col)
        Write_ICTVALUC(worksheet, rH, ITEM_CODE_col)

        WorkbookView1.ReleaseLock()

        Dim COLLECTION_CODE_ABSs As New List(Of String)
        For Each rowICTITEMX As DataRow In ASCDATA1.SelectDistinct("ICTITEMX", New String() {"COLLECTION_CODE"}).Select()
            Dim COLLECTION_CODE As String = (rowICTITEMX("COLLECTION_CODE") & "").Trim()
            Dim COLLECTION_CODE_ABS As String = ""

            If COLLECTION_CODE <> "" Then
                Dim r As DataRow = dst.Tables("SLTCOLL1").Rows.Find(COLLECTION_CODE)
                If r IsNot Nothing Then COLLECTION_CODE_ABS = r("COLLECTION_CODE") & ""
            End If
            If COLLECTION_CODE_ABS.Length > 6 Then
                Stop
                COLLECTION_CODE_ABS = Mid(COLLECTION_CODE_ABS, 1, 6)
            End If
            Dim rx As DataRow = dst.Tables("ICTCOLLX").NewRow
            rx("COLLECTION_CODE") = COLLECTION_CODE
            rx("COLLECTION_CODE_ABS") = COLLECTION_CODE_ABS
            dst.Tables("ICTCOLLX").Rows.Add(rx)

            If COLLECTION_CODE_ABS <> "" Then
                If COLLECTION_CODE_ABSs.Contains(COLLECTION_CODE_ABS) Then
                    dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTCOLLX", "COLLECTION_CODE_ABS",
                                                          $"Duplication: {COLLECTION_CODE_ABS} (raw='{COLLECTION_CODE}')"})
                Else
                    COLLECTION_CODE_ABSs.Add(COLLECTION_CODE_ABS)
                End If

                Dim row As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE_ABS)
                If row Is Nothing Then
                    row = dst.Tables("ICTCOLL1").NewRow
                    row("COLLECTION_CODE") = COLLECTION_CODE_ABS
                    dst.Tables("ICTCOLL1").Rows.Add(row)
                End If
                If (row("COLLECTION_NAME") & "") <> COLLECTION_CODE Then row("COLLECTION_NAME") = COLLECTION_CODE
            ElseIf COLLECTION_CODE <> "" Then
                dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTCOLL1", "COLLECTION_CODE",
                                                      $"Unmapped collection: {COLLECTION_CODE}"})
            End If
        Next
        Sort_grdColumns(grdICTCOLLX, "COLLECTION_CODE")
        Sort_grdColumns(grdICTCOLL1, "COLLECTION_CODE")

        Dim PROD_CODE_ABSs As New List(Of String)
        For Each rowICTITEMX As DataRow In ASCDATA1.SelectDistinct("ICTITEMX", New String() {"PROD_CODE"}).Select()
            Dim PROD_CODE As String = (rowICTITEMX("PROD_CODE") & "").Trim()
            Dim PROD_CODE_ABS As String = ""

            If PROD_CODE <> "" Then
                Select Case PROD_CODE.ToUpperInvariant()
                    Case "COLLATERAL" : PROD_CODE_ABS = "NCC"
                    Case "COMPONENT" : PROD_CODE_ABS = "COMP"
                    Case "CONSUMABLE" : PROD_CODE_ABS = "CNSMB"
                    Case "EXTERNAL" : PROD_CODE_ABS = "EXT"
                    Case "SELLABLE" : PROD_CODE_ABS = "SB"
                    Case Else
                        Dim r As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE)
                        If r IsNot Nothing Then PROD_CODE_ABS = r("PROD_CODE") & ""
                End Select
            End If

            Dim rx As DataRow = dst.Tables("ICTPRODX").NewRow
            rx("PROD_CODE") = PROD_CODE
            rx("PROD_CODE_ABS") = PROD_CODE_ABS
            dst.Tables("ICTPRODX").Rows.Add(rx)

            If PROD_CODE_ABS <> "" Then
                If PROD_CODE_ABSs.Contains(PROD_CODE_ABS) Then
                    dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTPRODX", "PROD_CODE_ABS",
                                                          $"Duplication: {PROD_CODE_ABS} (raw='{PROD_CODE}')"})
                Else
                    PROD_CODE_ABSs.Add(PROD_CODE_ABS)
                End If

                Dim row As DataRow = dst.Tables("ICTPROD1").Rows.Find(PROD_CODE_ABS)
                If row Is Nothing Then
                    row = dst.Tables("ICTPROD1").NewRow
                    row("PROD_CODE") = PROD_CODE_ABS
                    dst.Tables("ICTPROD1").Rows.Add(row)
                End If
                If (row("PROD_DESC") & "") <> PROD_CODE Then row("PROD_DESC") = PROD_CODE
            ElseIf PROD_CODE <> "" Then
                dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTPROD1", "PROD_CODE",
                                                      $"Unmapped product code: {PROD_CODE}"})
            End If
        Next
        Sort_grdColumns(grdICTPRODX, "PROD_CODE")
        Sort_grdColumns(grdICTPROD1, "PROD_CODE")

        Dim SEASON_CODE_ABSs As New List(Of String)
        For Each rowICTITEMX As DataRow In ASCDATA1.SelectDistinct("ICTITEMX", New String() {"SEASON_CODE"}).Select()
            Dim SEASON_CODE As String = (rowICTITEMX("SEASON_CODE") & "").Trim()
            Dim SEASON_CODE_ABS As String = ""

            If SEASON_CODE <> "" AndAlso Not {"TBC", "UNKNOWN"}.Contains(SEASON_CODE.ToUpperInvariant()) Then
                Dim parts() As String = SEASON_CODE.Split(" "c)
                If parts.Length >= 3 AndAlso IsNumeric(parts(0)) Then
                    Dim yearPart As String = parts(0)
                    Dim halfPart As String = parts(1).ToUpperInvariant()
                    If halfPart = "FIRST" Then
                        SEASON_CODE_ABS = yearPart & "S"
                    ElseIf halfPart = "SECOND" Then
                        SEASON_CODE_ABS = yearPart & "F"
                    End If
                End If
            End If

            Dim rx As DataRow = dst.Tables("ICTSEASX").NewRow
            rx("SEASON_CODE") = SEASON_CODE
            rx("SEASON_CODE_ABS") = SEASON_CODE_ABS
            dst.Tables("ICTSEASX").Rows.Add(rx)

            If SEASON_CODE_ABS <> "" Then
                If SEASON_CODE_ABSs.Contains(SEASON_CODE_ABS) Then
                    dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTSEASX", "SEASON_CODE_ABS",
                                                          $"Duplication: {SEASON_CODE_ABS} (raw='{SEASON_CODE}')"})
                Else
                    SEASON_CODE_ABSs.Add(SEASON_CODE_ABS)
                End If

                Dim row As DataRow = dst.Tables("ICTSEAS1").Rows.Find(SEASON_CODE_ABS)
                If row Is Nothing Then
                    row = dst.Tables("ICTSEAS1").NewRow
                    row("SEASON_CODE") = SEASON_CODE_ABS
                    dst.Tables("ICTSEAS1").Rows.Add(row)
                End If

                If (row("SEASON_DESC") & "") <> SEASON_CODE Then row("SEASON_DESC") = SEASON_CODE
                Dim y As String = Microsoft.VisualBasic.Mid(SEASON_CODE_ABS, 1, 4)
                If (row("SEASON_YEAR") & "") <> y Then row("SEASON_YEAR") = y
                Dim t As String = SEASON_CODE_ABS.Substring(SEASON_CODE_ABS.Length - 1, 1)
                If (row("SEASON_TYPE") & "") <> t Then row("SEASON_TYPE") = t
            ElseIf SEASON_CODE <> "" Then
                'dst.Tables("ARTSLPE1").Rows.Add(New String() {"ICTSEAS1", "SEASON_CODE",
                '                                      $"Unmapped season: {SEASON_CODE}"})
            End If
        Next


        Sort_grdColumns(grdICTSEASX, "SEASON_CODE")
        Sort_grdColumns(grdICTSEAS1, "SEASON_CODE")

        Dim ITEM_CLASS_CODE_ABSs As New List(Of String)
        For Each rowICTITEMX As DataRow In ASCDATA1.SelectDistinct("ICTITEMX", New String() {"ITEM_CLASS_CODE"}).Select()
            Dim ITEM_CLASS_CODE As String = (rowICTITEMX("ITEM_CLASS_CODE") & "").Trim()
            Dim ITEM_CLASS_CODE_ABS As String = ""

            Dim rx As DataRow = dst.Tables("ICTCLASX").NewRow
            rx("ITEM_CLASS_CODE") = ITEM_CLASS_CODE

            dst.Tables("ICTCLASX").Rows.Add(rx)

            If ITEM_CLASS_CODE <> "" Then
                Dim rowICTCLAS1 As DataRow = Nothing
                Dim row() As DataRow = dst.Tables("ICTCLAS1").Select($"ITEM_CLASS_DESC = '{ITEM_CLASS_CODE}'")
                If row.Length = 1 Then
                    rowICTCLAS1 = row(0)
                    ITEM_CLASS_CODE_ABS = rowICTCLAS1.Item("ITEM_CLASS_CODE")
                Else
                    ITEM_CLASS_CODE_ABS = dst.Tables("ICTCLAS1").Compute("MAX(ITEM_CLASS_CODE)", "") & ""
                    If ITEM_CLASS_CODE_ABS = "" Then
                        ITEM_CLASS_CODE_ABS = "001"
                    Else
                        ITEM_CLASS_CODE_ABS = Format(Val(ITEM_CLASS_CODE_ABS) + 1, "000")
                    End If

                    rowICTCLAS1 = dst.Tables("ICTCLAS1").NewRow
                    rowICTCLAS1("ITEM_CLASS_CODE") = ITEM_CLASS_CODE_ABS
                    rowICTCLAS1("ITEM_CLASS_DESC") = ITEM_CLASS_CODE
                    dst.Tables("ICTCLAS1").Rows.Add(rowICTCLAS1)
                End If
            End If

            rx("ITEM_CLASS_CODE_ABS") = ITEM_CLASS_CODE_ABS
        Next
        Sort_grdColumns(grdICTCLASX, "ITEM_CLASS_CODE")
        Sort_grdColumns(grdICTCLAS1, "ITEM_CLASS_CODE")

        Dim copyCols As New List(Of String)
        For Each col As String In ABS2SLP.Keys
            If dst.Tables("ICTITEMX").Columns.Contains(col) AndAlso dst.Tables("ICTITEM1").Columns.Contains(col) Then
                copyCols.Add(col)
            End If
        Next
        For Each rowICTITEMX As DataRow In dst.Tables("ICTITEMX").Select("")
            Dim ITEM_CODE As String = (rowICTITEMX("ITEM_CODE") & "").Trim()
            If ITEM_CODE = "" Then Continue For

            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            If rowICTITEM1 Is Nothing Then
                rowICTITEM1 = dst.Tables("ICTITEM1").NewRow()
                rowICTITEM1("ITEM_CODE") = ITEM_CODE
                dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
            End If
            With rowICTITEM1
                For Each C As String In copyCols
                    If String.Equals(C, "ITEM_CODE", StringComparison.OrdinalIgnoreCase) Then Continue For
                    Dim v As String = rowICTITEMX.Item(C) & ""
                    If v <> "" And v <> "-" And v <> "unknown" And v <> "TBC" Then
                        .Item(C) = rowICTITEMX.Item(C)
                    End If
                Next

                For Each dcol As DataColumn In dst.Tables("ICTITEMX").Columns
                    Dim C As String = dcol.ColumnName
                    Select Case C
                        Case "COLLECTION_CODE"
                            Dim COLLECTION_NAME As String = rowICTITEM1.Item("COLLECTION_CODE") & ""
                            If COLLECTION_NAME <> "" Then
                                Dim rowSLTCOLL1 As DataRow = dst.Tables("SLTCOLL1").Rows.Find(COLLECTION_NAME)
                                If rowSLTCOLL1 IsNot Nothing Then
                                    .Item("COLLECTION_CODE") = rowSLTCOLL1.Item("COLLECTION_CODE")
                                End If
                            End If
                        Case "ITEM_CATGY_CODE"
                            Dim ITEM_CATGY_CODE As String = rowICTITEMX("ITEM_CATGY_CODE") & ""
                            Dim mapped As String = Nothing
                            If ITEM_CATGY_CODE <> "" Then
                                Select Case ITEM_CATGY_CODE.ToUpperInvariant()
                                    Case "CORE" : mapped = "C"
                                    Case "LIMITED EDITION", "LIMITED-EDITION", "LIMITED" : mapped = "E"
                                    Case "UNKNOWN" : mapped = "I"
                                End Select
                            End If

                            If mapped IsNot Nothing Then
                                .Item("ITEM_CATGY_CODE") = mapped
                            End If
                        Case "ITEM_CLASS_CODE"
                            Dim ITEM_CLASS_CODE As String = rowICTITEMX("ITEM_CLASS_CODE") & ""
                            If ITEM_CLASS_CODE & "" <> "" Then
                                Dim mapped As DataRow = dst.Tables("ICTCLASX").Rows.Find(ITEM_CLASS_CODE)
                                If mapped IsNot Nothing AndAlso mapped("ITEM_CLASS_CODE_ABS") <> "" Then
                                    .Item("ITEM_CLASS_CODE") = mapped("ITEM_CLASS_CODE_ABS")
                                End If
                            End If

                        Case "ITEM_STATUS"
                            Dim ITEM_STATUS As String = rowICTITEMX("ITEM_STATUS") & ""
                            Dim mapped As String = "I"
                            If ITEM_STATUS.Equals("ACTIVE", StringComparison.OrdinalIgnoreCase) Then
                                mapped = "A"
                            End If

                            .Item("ITEM_STATUS") = mapped

                        Case "SEASON_CODE"
                            Dim rawSeason As String = (rowICTITEMX("SEASON_CODE") & "").Trim()
                            Select Case True
                                Case rawSeason = "" _
                                     OrElse rawSeason.Equals("TBC", StringComparison.OrdinalIgnoreCase) _
                                     OrElse rawSeason.Equals("UNKNOWN", StringComparison.OrdinalIgnoreCase)
                                    .Item("SEASON_CODE") = DBNull.Value

                                Case Else
                                    Dim parts() As String = rawSeason.Split(" "c)
                                    Dim seasonABS As String = ""
                                    If parts.Length >= 2 AndAlso IsNumeric(parts(0)) Then
                                        Dim yr As String = parts(0)
                                        Dim half As String = parts(1).ToUpperInvariant()
                                        If half = "FIRST" Then seasonABS = yr & "S"
                                        If half = "SECOND" Then seasonABS = yr & "F"
                                    End If
                                    .Item("SEASON_CODE") = If(seasonABS = "", CType(DBNull.Value, Object), seasonABS)

                            End Select

                    End Select
                Next
                'dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
            End With
        Next

        Dim eMessage As String = clsASCBASE1.EnforceConstraints(True)
        If eMessage <> "" Then
            Dim TABLE_NAME_ERROR As String = "*"
            Dim COLUMN_NAME_ERROR As String = "*"
            For Each MSG As String In Split(eMessage, vbCrLf)
                If MSG <> "" Then
                    If MSG.StartsWith("Table: ") Then
                        TABLE_NAME_ERROR = Mid(MSG, 8)
                    Else
                        Dim lookfor As String = " Row Error: Column 'C"
                        If MSG.Contains(lookfor) Then '  MSG.StartsWith(" Row Error: Column 'C") Then
                            Dim i2 As Integer = InStr(MSG, lookfor)
                            COLUMN_NAME_ERROR = Mid(MSG, 21 + i2 - 1)
                            Dim i As Integer = InStr(COLUMN_NAME_ERROR, "'")
                            If i <> 0 Then COLUMN_NAME_ERROR = Mid(COLUMN_NAME_ERROR, 1, i - 1)
                        End If
                        dst.Tables("ARTSLPE1").Rows.Add(TABLE_NAME_ERROR, COLUMN_NAME_ERROR, MSG)
                        TABLE_NAME_ERROR = "*"
                        COLUMN_NAME_ERROR = "*"
                    End If
                End If
            Next
        Else
            EnforceConstraints(True)
        End If

        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Database")

        BeginTrans()

        For Each TABLE_NAME In ABS_TABLES
            ASCMAIN1.Progress("-", TABLE_NAME)
            Update_Record_TDA(TABLE_NAME)
        Next

        Update_Record_TDA("ICTRETLC")
        Update_Record_TDA("ICTVALUC")

        ASCMAIN1.sql = "update ARTCUST1 set CUST_ECOM_IND = NULL, CUST_PII_IND=NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "update ARTCUST1 set CUST_ECOM_IND = '1', CUST_PII_IND='1' where cust_code in
( 'AAFESCOM'
, 'JOY'
, 'NORDDROP'
, 'QVC'
, 'SLIPCOMC'
, 'SLIPCOMU'
, 'WAYFAIR'
, 'WESTELM'
, 'POOSH'
, 'POTTERBAR'
, 'SALONCENTRICDS')"
        ASCDATA1.ExecuteSQL()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Update Complete")

        Archive_Input_File()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTITEMX, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        'If tlb_pop.Tools.Exists("Include Inactive") Then
        'End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"
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
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        'WorkbookView1.GetLock()
        'Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("RSFSSPL1.XLSX_NO") & ".XLSX"
        'WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        'Show_Document(FILENAME)
        'WorkbookView1.ReleaseLock()
    End Sub
    'Private Sub Link_Dropbox() 'Opens dropbox in browser, paste the code from URL bar into prompt
    '    Dim url = DropboxAuth.BuildAuthUrl()
    '    Process.Start(url)

    '    Dim code As String = InputBox("After you sign in, copy the 'code' value from the URL bar and paste it here:", "Paste Dropbox code")
    '    If String.IsNullOrWhiteSpace(code) Then
    '        MsgBox("No code pasted. Linking canceled.")
    '        Exit Sub
    '    End If

    '    Try
    '        Dim rt As String = DropboxTokenExchange.ExchangeCode(code.Trim(), DropboxAuth.CodeVerifier)
    '        If String.IsNullOrWhiteSpace(rt) Then
    '            MsgBox("No refresh token returned. Check app settings / code pasted.", MsgBoxStyle.Exclamation)
    '            Exit Sub
    '        End If
    '        TokenStore.SaveRefreshToken(rt)
    '        Dim rtEsc As String = rt.Replace("'", "''")
    '        ASCMAIN1.sql = $"UPDATE RSTPARM1 SET RS_PARM_DBOX_REF_TOKEN = '{rtEsc}' WHERE RS_PARM_KEY = 'Z'"
    '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
    '        MsgBox("Dropbox linked successfully.", MsgBoxStyle.Information)
    '    Catch ex As Exception
    '        MsgBox("Linking failed: " & ex.Message, MsgBoxStyle.Critical)
    '    End Try
    'End Sub
    ' Turn a normal Dropbox share URL into a direct-download URL
    'Private Function ToDirectDownloadUrl(sharedUrl As String) As String
    '    If String.IsNullOrWhiteSpace(sharedUrl) Then Return sharedUrl
    '    Dim u = sharedUrl.Replace("www.dropbox.com", "dl.dropboxusercontent.com")
    '    If u.Contains("?") Then
    '        If Regex.IsMatch(u, "(\?|&)dl=\d") Then
    '            u = Regex.Replace(u, "(\?|&)dl=\d", "$1dl=1")
    '        Else
    '            u &= "&dl=1"
    '        End If
    '    Else
    '        u &= "?dl=1"
    '    End If
    '    Return u
    'End Function

    ' Exchange refresh_token -> short-lived access_token to call Dropbox API
    'Private Function GetAccessTokenFromRefresh() As String
    '    Dim REFRESH_TOKEN As String = ASCDATA1.GetDataValue("SELECT RS_PARM_DBOX_REF_TOKEN FROM RSTPARM1 WHERE RS_PARM_KEY = 'Z'") & ""
    '    Using wc As New WebClient()
    '        Dim data As New NameValueCollection() From {
    '        {"grant_type", "refresh_token"},
    '        {"refresh_token", REFRESH_TOKEN},
    '        {"client_id", DropboxConfig.AppKey}
    '    }
    '        Dim bytes = wc.UploadValues("https://api.dropboxapi.com/oauth2/token", "POST", data)
    '        Dim json = Encoding.UTF8.GetString(bytes)
    '        Dim jt = JObject.Parse(json)
    '        Return jt.Value(Of String)("access_token")
    '    End Using
    'End Function

    ' Download bytes from a shared link; handles public and passworded links
    'Private Function DownloadSharedLinkToBytes(sharedUrl As String, Optional linkPassword As String = Nothing) As Byte()
    '    ' Try simple direct download first
    '    Try
    '        Using wc As New WebClient()
    '            Return wc.DownloadData(ToDirectDownloadUrl(sharedUrl))
    '        End Using
    '    Catch
    '    End Try

    '    Dim access = GetAccessTokenFromRefresh()
    '    Dim apiArg As New JObject From {{"url", sharedUrl}}
    '    If Not String.IsNullOrEmpty(linkPassword) Then apiArg("link_password") = linkPassword

    '    Using wc As New WebClient()
    '        wc.Headers(HttpRequestHeader.Authorization) = "Bearer " & access
    '        wc.Headers.Add("Dropbox-API-Arg", apiArg.ToString())
    '        Return wc.UploadData("https://content.dropboxapi.com/2/sharing/get_shared_link_file", "POST", New Byte() {})
    '    End Using
    'End Function

    '' Ask for a link, download, and open in SpreadsheetGear
    'Private Sub LoadFromDropboxSharedLink()
    '    Dim url As String = ASCDATA1.GetDataValue("SELECT RS_PARM_DBOX_ITEM_FILE FROM RSTPARM1 WHERE RS_PARM_KEY = 'Z'")
    '    If String.IsNullOrWhiteSpace(url) Then Exit Sub

    '    Dim bytes As Byte()
    '    Try
    '        bytes = DownloadSharedLinkToBytes(url)
    '    Catch ex As Exception
    '        MsgBox("Download failed: " & ex.Message, MsgBoxStyle.Critical)
    '        Exit Sub
    '    End Try

    '    Try
    '        Using ms As New MemoryStream(bytes)
    '            WorkbookView1.GetLock()
    '            ' open the workbook from bytes
    '            WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbookSet().Workbooks.OpenFromStream(ms)
    '            WorkbookView1.ReleaseLock()
    '        End Using
    '        MsgBox("Loaded workbook from Dropbox.", MsgBoxStyle.Information)
    '    Catch ex As Exception
    '        MsgBox("Failed to open workbook: " & ex.Message, MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    Private Function Col_To_Index(colLetters As String) As Integer
        colLetters = colLetters.Trim().ToUpperInvariant()
        Dim n As Integer = 0
        For i As Integer = 0 To colLetters.Length - 1
            n = n * 26 + (Asc(colLetters(i)) - Asc("A"c) + 1)
        Next
        Return n - 1
    End Function
    Private Function To_Dec_Or_Null(v As Object) As Decimal?
        If v Is Nothing Then Return Nothing
        Dim s As String = (v & "").Trim()
        If s = "" OrElse s = "-" OrElse s.Equals("TBC", StringComparison.OrdinalIgnoreCase) Then Return Nothing
        s = s.Replace(",", "")
        Dim d As Decimal
        If Decimal.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then
            If d = 0D Then Return Nothing
            Return d
        End If
        Return Nothing
    End Function
    Private Sub Write_ICTRETLC(ws As SpreadsheetGear.IWorksheet, headerRow As Integer, itemCodeColIndex As Integer)
        Dim OPS_US_CURR As String = "202508"  ' US current as of 08/01/2025
        Dim OPS_CA_CURR As String = "202510"  ' CA current as of 10/01/2025
        Dim OPS_POST As String = "202502"  ' 02/01/2025
        Dim OPS_PRE As String = "202001"  ' pre 02/01/2025

        Dim colUS_Current As Integer = Col_To_Index("FU") ' current US
        Dim colCA_Current As Integer = Col_To_Index("FV") ' current CA
        Dim colUS_Post As Integer = Col_To_Index("HB") ' 02/01/2025 US
        Dim colCA_Post As Integer = Col_To_Index("HC") ' 02/01/2025 CA
        Dim colUS_Pre As Integer = Col_To_Index("II") ' pre 02/01/2025 US
        Dim colCA_Pre As Integer = Col_To_Index("IJ") ' pre 02/01/2025 CA

        Dim t As DataTable = dst.Tables("ICTRETLC")
        t.Rows.Clear()

        Dim used As SpreadsheetGear.IRange = ws.UsedRange
        Dim lastRow As Integer = used.Row + used.RowCount - 1
        For r As Integer = headerRow + 1 To lastRow
            Dim a As String = ws.Cells(r, 0).Value & ""
            Dim b As String = ws.Cells(r, 1).Value & ""
            If a = "" AndAlso b = "" Then Exit For

            Dim itemCode As String = (ws.Cells(r, itemCodeColIndex).Value & "").Trim()
            If itemCode = "" Then Continue For

            ' --- US CURRENT (202508)
            Dim vUS_Cur As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Current).Value)
            If vUS_Cur.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_US_CURR
                dr("ITEM_RETAIL_PRICE") = vUS_Cur.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- CA CURRENT (202510)
            Dim vCA_Cur As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Current).Value)
            If vCA_Cur.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_CA_CURR
                dr("ITEM_RETAIL_PRICE") = vCA_Cur.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- POST 02/01/2025 (202502)
            Dim vUS_Post As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Post).Value)
            If vUS_Post.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_POST
                dr("ITEM_RETAIL_PRICE") = vUS_Post.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            Dim vCA_Post As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Post).Value)
            If vCA_Post.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_POST
                dr("ITEM_RETAIL_PRICE") = vCA_Post.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- PRE 02/01/2025 (202001)
            Dim vUS_Pre As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Pre).Value)
            If vUS_Pre.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_PRE
                dr("ITEM_RETAIL_PRICE") = vUS_Pre.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            Dim vCA_Pre As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Pre).Value)
            If vCA_Pre.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_PRE
                dr("ITEM_RETAIL_PRICE") = vCA_Pre.Value
                dr("ITEM_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If
        Next
    End Sub
    Private Sub Write_ICTVALUC(ws As SpreadsheetGear.IWorksheet, headerRow As Integer, itemCodeColIndex As Integer)
        ' Periods
        Dim OPS_US_CURR As String = "202508"  ' US current as of 08/01/2025
        Dim OPS_CA_CURR As String = "202510"  ' CA current as of 10/01/2025
        Dim OPS_POST As String = "202502"   ' 02/01/2025
        Dim OPS_PRE As String = "202001"   ' pre 02/01/2025

        Dim colUS_Current As Integer = Col_To_Index("GS") ' current US value
        Dim colCA_Current As Integer = Col_To_Index("GT") ' current CA value
        Dim colUS_Post As Integer = Col_To_Index("HZ") ' 02/01/2025 US value
        Dim colCA_Post As Integer = Col_To_Index("IA") ' 02/01/2025 CA value
        Dim colUS_Pre As Integer = Col_To_Index("JG") ' pre 02/01/2025 US value
        Dim colCA_Pre As Integer = Col_To_Index("JH") ' pre 02/01/2025 CA value

        Dim t As DataTable = dst.Tables("ICTVALUC")
        t.Rows.Clear()

        Dim used As SpreadsheetGear.IRange = ws.UsedRange
        Dim lastRow As Integer = used.Row + used.RowCount - 1
        For r As Integer = headerRow + 1 To lastRow
            Dim a As String = ws.Cells(r, 0).Value & ""
            Dim b As String = ws.Cells(r, 1).Value & ""
            If a = "" AndAlso b = "" Then Exit For

            Dim itemCode As String = (ws.Cells(r, itemCodeColIndex).Value & "").Trim()
            If itemCode = "" Then Continue For

            ' --- US CURRENT (202508)
            Dim vUS_Cur As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Current).Value)
            If vUS_Cur.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_US_CURR
                dr("ITEM_RETAIL_VALUE_PRICE") = vUS_Cur.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- CA CURRENT (202510)
            Dim vCA_Cur As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Current).Value)
            If vCA_Cur.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_CA_CURR
                dr("ITEM_RETAIL_VALUE_PRICE") = vCA_Cur.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- POST 02/01/2025 (202502)
            Dim vUS_Post As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Post).Value)
            If vUS_Post.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_POST
                dr("ITEM_RETAIL_VALUE_PRICE") = vUS_Post.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            Dim vCA_Post As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Post).Value)
            If vCA_Post.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_POST
                dr("ITEM_RETAIL_VALUE_PRICE") = vCA_Post.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            ' --- PRE 02/01/2025 (202001)
            Dim vUS_Pre As Decimal? = To_Dec_Or_Null(ws.Cells(r, colUS_Pre).Value)
            If vUS_Pre.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "USD"
                dr("OPS_YYYYPP") = OPS_PRE
                dr("ITEM_RETAIL_VALUE_PRICE") = vUS_Pre.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If

            Dim vCA_Pre As Decimal? = To_Dec_Or_Null(ws.Cells(r, colCA_Pre).Value)
            If vCA_Pre.HasValue Then
                Dim dr As DataRow = t.NewRow()
                dr("ITEM_CODE") = itemCode
                dr("CURR_CODE") = "CAD"
                dr("OPS_YYYYPP") = OPS_PRE
                dr("ITEM_RETAIL_VALUE_PRICE") = vCA_Pre.Value
                dr("ITEM_VALUE_PRICE") = DBNull.Value
                dr("ITEM_CATGY_CODE") = DBNull.Value
                t.Rows.Add(dr)
            End If
        Next
    End Sub
    Private Function Get_Unique_Path(sourceName As String) As String
        Dim name As String = If(String.IsNullOrWhiteSpace(sourceName), "ItemMaster.xlsx", sourceName)
        Dim baseName As String = Path.GetFileNameWithoutExtension(name)
        Dim ext As String = Path.GetExtension(name)
        If String.IsNullOrEmpty(ext) Then ext = ".xlsx"

        Dim ts As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim candidate As String = Path.Combine(ARCHIVE_DIR, $"{baseName}_{ts}{ext}")

        Dim i As Integer = 1
        While File.Exists(candidate)
            candidate = Path.Combine(ARCHIVE_DIR, $"{baseName}_{ts}_{i}{ext}")
            i += 1
        End While
        Return candidate
    End Function
    Private Sub Archive_Input_File()
        Try
            If Not Directory.Exists(ARCHIVE_DIR) Then
                Directory.CreateDirectory(ARCHIVE_DIR)
            End If

            If Not String.IsNullOrWhiteSpace(FILENAME) AndAlso File.Exists(FILENAME) Then
                Dim dest As String = Get_Unique_Path(Path.GetFileName(FILENAME))
                File.Copy(FILENAME, dest, overwrite:=False)
                ASCMAIN1.Progress($"Archived source file: {dest}")
            Else
                Dim wbName As String = TryCast(WorkbookView1?.ActiveWorkbook?.Name, String)
                If String.IsNullOrWhiteSpace(wbName) Then wbName = "ItemMaster.xlsx"
                Dim dest As String = Get_Unique_Path(wbName)

                Dim fmt As SpreadsheetGear.FileFormat =
                    If(String.Equals(Path.GetExtension(dest), ".xls", StringComparison.OrdinalIgnoreCase),
                       SpreadsheetGear.FileFormat.Excel8,
                       SpreadsheetGear.FileFormat.OpenXMLWorkbook)

                WorkbookView1.GetLock()
                Try
                    WorkbookView1.ActiveWorkbook.SaveAs(dest, fmt)
                Finally
                    WorkbookView1.ReleaseLock()
                End Try
                ASCMAIN1.Progress($"Archived workbook: {dest}")
            End If
        Catch ex As Exception
            ASCMAIN1.Progress("Archive step failed: " & ex.Message)
        End Try
    End Sub
    Private Function Normalize_ABS(s As String) As String
        If s Is Nothing Then Return ""
        s = s.Trim().ToUpperInvariant()
        If s.Length = 0 OrElse s.Length > 6 Then Return ""
        Return s
    End Function
    Private Function Is_Unique(t As DataTable, pkField As String, code As String) As Boolean
        If String.IsNullOrEmpty(code) Then Return False
        If t.PrimaryKey IsNot Nothing AndAlso t.PrimaryKey.Length > 0 AndAlso
           t.PrimaryKey(0).ColumnName.Equals(pkField, StringComparison.OrdinalIgnoreCase) Then
            Return (t.Rows.Find(code) Is Nothing)
        End If
        Return t.Select($"{pkField} = '{code.Replace("'", "''")}'").Length = 0
    End Function
    Private Sub Add_ICTCOLL1_and_ICTCOLL0(absCode As String, Optional displayName As String = "")
        Dim t As DataTable = dst.Tables("ICTCOLL1")
        If t Is Nothing OrElse String.IsNullOrEmpty(absCode) Then Exit Sub
        If t.Rows.Find(absCode) IsNot Nothing Then Exit Sub

        Dim n As DataRow = t.NewRow()
        If t.Columns.Contains("COLLECTION_CODE") Then n("COLLECTION_CODE") = absCode
        If t.Columns.Contains("COLLECTION_NAME") AndAlso displayName <> "" Then n("COLLECTION_NAME") = displayName
        If t.Columns.Contains("HC_CODE") Then n("HC_CODE") = absCode
        If t.Columns.Contains("BRAND_CODE") Then n("BRAND_CODE") = "SLIP"

        t.Rows.Add(n)

        Dim c0 = dst.Tables("ICTCOLL0")
        If c0 IsNot Nothing AndAlso c0.Rows.Find(absCode) Is Nothing Then
            Dim nr = c0.NewRow()
            If c0.Columns.Contains("HC_CODE") Then nr("HC_CODE") = absCode
            c0.Rows.Add(nr)
        End If
    End Sub
    Private Sub Add_ICTPROD1(absCode As String, Optional srcName As String = "")
        Dim t As DataTable = dst.Tables("ICTPROD1")
        If t Is Nothing OrElse String.IsNullOrEmpty(absCode) Then Exit Sub
        If t.Rows.Find(absCode) IsNot Nothing Then Exit Sub

        Dim n As DataRow = t.NewRow()
        If t.Columns.Contains("PROD_CODE") Then n("PROD_CODE") = absCode
        If t.Columns.Contains("PROD_DESC") AndAlso srcName <> "" Then n("PROD_DESC") = srcName
        t.Rows.Add(n)
    End Sub
    Private Sub ReEnterEdit(cell As UltraWinGrid.UltraGridCell, message As String)
        MsgBox(message, MsgBoxStyle.Exclamation, "Invalid Code")
        cell.Value = ""
        cell.Row.Update()
    End Sub
    Private Sub grdICTCOLLX_BeforeCellUpdate(sender As Object, e As UltraWinGrid.BeforeCellUpdateEventArgs) _
    Handles grdICTCOLLX.BeforeCellUpdate

        If e.Cell.Column.Key <> "COLLECTION_CODE_ABS" Then Exit Sub

        Dim raw As String = (e.NewValue & "").Trim().ToUpperInvariant()
        If raw.Length = 0 OrElse raw.Length > 6 Then
            MsgBox("Collection Code ABS must be 1–6 characters.", MsgBoxStyle.Exclamation, "Invalid Code")
            e.Cancel = True
            Exit Sub
        End If

        Dim tMaster As DataTable = dst.Tables("ICTCOLL1")
        If Not Is_Unique(tMaster, "COLLECTION_CODE", raw) Then
            MsgBox($"Collection Code ABS '{raw}' already exists. Enter a unique code.", MsgBoxStyle.Exclamation, "Invalid Code")
            e.Cancel = True
            Exit Sub
        End If

    End Sub
    Private Sub grdICTCOLLX_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) _
    Handles grdICTCOLLX.AfterCellUpdate

        If e.Cell.Column.Key <> "COLLECTION_CODE_ABS" Then Exit Sub

        Dim absCode As String = (e.Cell.Text & "").Trim().ToUpperInvariant()
        If absCode = "" Then Exit Sub

        Dim displayName As String = (e.Cell.Row.Cells("COLLECTION_CODE").Text & "")
        Add_ICTCOLL1_and_ICTCOLL0(absCode, displayName)
    End Sub
    Private Sub grdICTPRODX_BeforeCellUpdate(sender As Object, e As UltraWinGrid.BeforeCellUpdateEventArgs) _
    Handles grdICTPRODX.BeforeCellUpdate

        If e.Cell.Column.Key <> "PROD_CODE_ABS" Then Exit Sub

        Dim raw As String = (e.NewValue & "").Trim().ToUpperInvariant()

        If raw.Length = 0 OrElse raw.Length > 6 Then
            MsgBox("Product Code ABS must be 1–6 characters.", MsgBoxStyle.Exclamation, "Invalid Code")
            e.Cancel = True
            Exit Sub
        End If

        Dim tMaster As DataTable = dst.Tables("ICTPROD1")
        If Not Is_Unique(tMaster, "PROD_CODE", raw) Then
            MsgBox($"Product Code ABS '{raw}' already exists. Enter a unique code.", MsgBoxStyle.Exclamation, "Invalid Code")
            e.Cancel = True
            Exit Sub
        End If
    End Sub
    Private Sub grdICTPRODX_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) _
    Handles grdICTPRODX.AfterCellUpdate

        If e.Cell.Column.Key <> "PROD_CODE_ABS" Then Exit Sub

        Dim absCode As String = (e.Cell.Text & "").Trim().ToUpperInvariant()
        If absCode = "" Then Exit Sub

        Dim srcName As String = (e.Cell.Row.Cells("PROD_CODE").Text & "")
        Add_ICTPROD1(absCode, srcName)
    End Sub


End Class

'Public Module DropboxConfig 'Holds the AppKey, our RedirectUri, and Token File that saves the refresh token
'    Public Const AppKey As String = "howukkwaelbwi8w"
'    Public Const RedirectUri As String = "http://localhost:53682/authorize-finish"
'    Public ReadOnly TokenFile As String = IO.Path.Combine(
'        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
'        "Slip", "Importer", "dropbox.refresh.token")
'End Module

'Public Module TokenStore 'Encrypts the token bytes and saves to TokenFile using DPAPI
'    Public Sub SaveRefreshToken(refreshToken As String)
'        Dim bytes = Encoding.UTF8.GetBytes(refreshToken)
'        Dim enc = ProtectedData.Protect(bytes, Nothing, DataProtectionScope.CurrentUser)
'        Dim dir = IO.Path.GetDirectoryName(DropboxConfig.TokenFile)
'        If Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
'        IO.File.WriteAllBytes(DropboxConfig.TokenFile, enc)
'    End Sub

'    Public Function LoadRefreshToken() As String 'Reads the bytes, decrypts with DPAPI and returns the plain refresh token string
'        If Not IO.File.Exists(DropboxConfig.TokenFile) Then Return Nothing
'        Dim enc = IO.File.ReadAllBytes(DropboxConfig.TokenFile)
'        Dim dec = ProtectedData.Unprotect(enc, Nothing, DataProtectionScope.CurrentUser)
'        Return Encoding.UTF8.GetString(dec)
'    End Function
''End Module
'Public Module DropboxAuth 'Implements OAuth 2.0 with PKCE, gets one time code from Dropbox
'    Private _codeVerifier As String

'    Private Function B64Url(bytes As Byte()) As String
'        Return Convert.ToBase64String(bytes).TrimEnd("="c).Replace("+", "-").Replace("/", "_")
'    End Function

'    Private Function NewRandom(len As Integer) As String
'        Using rng As New RNGCryptoServiceProvider()
'            Dim b(len - 1) As Byte : rng.GetBytes(b)
'            Return B64Url(b)
'        End Using
'    End Function

'    Public Function BuildAuthUrl() As String
'        _codeVerifier = NewRandom(64)

'        Dim codeChallenge As String
'        Using sha As SHA256 = SHA256.Create()
'            Dim hash = sha.ComputeHash(Encoding.ASCII.GetBytes(_codeVerifier))
'            codeChallenge = B64Url(hash)
'        End Using

'        Dim scope As String = "files.content.read files.metadata.read sharing.read"
'        Dim qs As String =
'            "response_type=code" &
'            "&client_id=" & Uri.EscapeDataString(DropboxConfig.AppKey) &
'            "&redirect_uri=" & Uri.EscapeDataString(DropboxConfig.RedirectUri) &
'            "&code_challenge=" & Uri.EscapeDataString(codeChallenge) &
'            "&code_challenge_method=S256" &
'            "&token_access_type=offline" &
'            "&scope=" & Uri.EscapeDataString(scope)

'        Return "https://www.dropbox.com/oauth2/authorize?" & qs
'    End Function

'    Public ReadOnly Property CodeVerifier As String
'        Get
'            Return _codeVerifier
'        End Get
'    End Property
'End Module
'Public Module DropboxTokenExchange 'takes the one-time code plus verifier and sends to Dropbox, Dropbox replies with refresh token
'    Public Function ExchangeCode(authCode As String, codeVerifier As String) As String
'        Using wc As New WebClient()
'            Dim data As New NameValueCollection() From {
'                {"code", authCode},
'                {"grant_type", "authorization_code"},
'                {"client_id", DropboxConfig.AppKey},
'                {"code_verifier", codeVerifier},
'                {"redirect_uri", DropboxConfig.RedirectUri}
'            }
'            Dim bytes = wc.UploadValues("https://api.dropboxapi.com/oauth2/token", "POST", data)
'            Dim json = Encoding.UTF8.GetString(bytes)
'            Dim jt = JObject.Parse(json)
'            Return jt.Value(Of String)("refresh_token")
'        End Using
'    End Function
'End Module
