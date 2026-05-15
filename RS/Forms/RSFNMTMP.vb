Public Class RSFNMTMP
    Dim EXCEL_SHEET As String
    Dim objConnection As System.Data.OleDb.OleDbConnection
    Dim FILENAME As String = ""
    Dim sqlEDT852TX As String
    Dim EDT852TX As String
    Dim OPS_YYYYWW As String
    Dim CUST_CODE As String
    Dim rowGLTPARM3 As DataRow
    Dim OPS_YYYYPP As String
    Dim WEEK_END_DATE As Date
    Dim BAD_ITEMS_CHECKED As Boolean

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYW("OPS_YYYYWW", ASCMAIN1.CYW, -52, 0, -1)
        'Sort_cmbColumns(Absx1.cmbFor("OPS_YYYYWW"), "YYYYWW".ToLower)

        With dst

            sqlEDT852TX = "Select EDT852T1.EDI_DOC_SEQ_NO, EDT852T1.EDI_DEPT_NO" & vbCrLf _
            & ", EDT852T1.OPS_YYYYWW, EDT852T1.CUST_CODE, EDT852T1.COLLECTION_CODE" & vbCrLf _
            & " from EDT852T1" & vbCrLf _
            & " where EDT852T1.CUST_CODE = :PARM1" & vbCrLf _
            & "   and EDT852T1.OPS_YYYYWW = :PARM2" & vbCrLf _
            & "   and EDT852T1.EDI_STATUS = 'M'" & vbCrLf
            EDT852TX = ASCMAIN1.Temp_Table(Replace(Replace(sqlEDT852TX, ":PARM1", "Null"), ":PARM2", "Null"))
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TX & " Add Primary Key (EDI_DOC_SEQ_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TX & " Add QTY_SOLD NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TX & " Add AMT_SOLD NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TX & " Add QTY_EOW NUMBER (8,0)")
            ASCMAIN1.sql = "Select * from " & EDT852TX & " EDT852TX"
            Create_TDA(.Tables.Add, "EDT852TX", "**", 0, False)
            .Tables("EDT852TX").Columns.Add("SELECTED", GetType(System.String))

            Create_TDA(.Tables.Add, "RSTNMTMP", "*")
            .Tables("RSTNMTMP").Columns.Add("BAD_ITEM", GetType(System.String))
            .Tables("RSTNMTMP").Columns.Add("COLLECTION_CODE", GetType(System.String))
            .Tables("RSTNMTMP").Columns.Add("ITEM_RETAIL_PRICE", GetType(System.Decimal))
            .Tables("RSTNMTMP").Columns.Add("AMT_SOLD", GetType(System.Decimal), "ISNULL(QTY_SOLD,0) * ISNULL(ITEM_RETAIL_PRICE,0)")

            Create_TDA(.Tables.Add, "RSTITEMX", "*", 1)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)

            Create_TDA(.Tables.Add, "RSTRETL1", "*", 0)
            Create_TDA(.Tables.Add, "EDT852T1", "*", 0)
        End With

        grdRSTNMTMP.DataSource = dst.Tables("RSTNMTMP")
        grdEDT852TX.DataSource = dst.Tables("EDT852TX")

        Create_Summary(grdRSTNMTMP, "DEPT", "Count")
        Create_Summary(grdRSTNMTMP, "QTY_SOLD")
        Create_Summary(grdRSTNMTMP, "QTY_OH")

        With grdRSTNMTMP.DisplayLayout.Bands("RSTNMTMP")
            .Columns("CLS").Header.Fixed = True
            .Columns("DEPT").Header.Fixed = True
            .Columns("STORE").Header.Fixed = True
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTNMTMP.DisplayLayout.Bands(0).Columns
            If gcol.Key = "ITEM_CODE" Or gcol.Key = "CUST_STORE_NO" Then
                gcol.CellAppearance.BackColor = Color.LightPink
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If

            If gcol.Key = "CUST_CODE" _
            Or gcol.Key = "OPS_YYYYWW" _
            Or gcol.Key = "XNO" _
            Or gcol.Key = "COLLECTION_CODE" _
            Or gcol.Key = "ITEM_RETAIL_PRICE" _
            Or gcol.Key = "AMT_SOLD" _
            Or gcol.Key = "BAD_ITEM" Then
                gcol.Hidden = True
            End If
        Next

        For Each gcol As UltraWinGrid.UltraGridColumn In grdEDT852TX.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SELECTED" Then
                gcol.CellAppearance.BackColor = Color.LightPink
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "No Customer Specified"
                Else
                    Validate_Code("CUST_CODE")
                End If

                If Absx1.cmbFor("OPS_YYYYWW").Text = "" Then
                    EMsg &= vbCr & "No Week Specified"
                End If

            Case "Load Spreadsheet"

                FILENAME = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls"
                    openFileDialog1.RestoreDirectory = True

                    'Excel_Import = -1

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using
                ' GC.Collect()

                If FILENAME <> "" Then
                    Try
                        Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                            "data source=" & FILENAME & ";" & _
                            "Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
                        Dim SHEETS As Int32 = 0
                        Dim dbSchema As DataTable
                        Using objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                            objConnection.Open()
                            dbSchema = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                            If dbSchema.Rows.Count > 0 Then
                                EXCEL_SHEET = dbSchema.Rows(0).Item("TABLE_NAME")
                            End If
                            SHEETS = dbSchema.Rows.Count
                        End Using

                        If SHEETS = 0 Then
                            MsgBox("No Sheets Found")
                            Exit Sub
                        End If

                        If SHEETS > 1 Then
                            Dim dtx As New DataTable
                            dtx.Columns.Add("TABLE_NAME")
                            Dim TABLE_NAME As String = ""
                            For Each row As DataRow In dbSchema.Rows

                                Dim SHEET_NAME As String = row.Item("TABLE_NAME")
                                If SHEET_NAME.EndsWith("$") Or (SHEET_NAME.StartsWith("'") And SHEET_NAME.EndsWith("$'")) Then
                                    If SHEET_NAME.StartsWith("'") Then
                                        SHEET_NAME = Mid(SHEET_NAME, 2, Len(SHEET_NAME) - 2)
                                    End If
                                    TABLE_NAME = Mid(SHEET_NAME, 1, Len(SHEET_NAME) - 1)
                                    dtx.Rows.Add(TABLE_NAME)
                                End If
                            Next
                            If dtx.Rows.Count = 1 Then
                                EXCEL_SHEET = TABLE_NAME & "$"
                            Else
                                Using frmmsg As New ASFMSGBF
                                    frmmsg.Show_grd(dtx, Me, "Select Excel Sheet to Load")
                                    If frmmsg.grow Is Nothing Then
                                        Exit Sub
                                    End If
                                    EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text & "$"
                                End Using
                            End If
                        Else

                        End If


                    Catch ex As Exception
                        MsgBox("Exception Occurred:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error Opening Excel Workbook")
                    Finally

                    End Try

                Else
                    EMsg &= "No File Selected"
                End If

            Case "Update"

                Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", OPS_YYYYWW)

                For Each row As DataRow In ASCDATA1.SelectDistinct("RSTNMTMP", "WE_DATE").Rows
                    Dim WE_DATE As Date = row.Item(0)
                    If Format(WE_DATE, "yyyyMMdd") <> Format(rowGLTPARM3.Item("WEEK_END_DATE"), "yyyyMMdd") Then
                        EMsg &= vbCr & "W/E date " & Format(WE_DATE, "MM/dd/yyyy") & " in grid does not agree with W/E date for Week " & OPS_YYYYWW & " (" & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yyyy") & ")"
                    End If
                Next

                Check_Items(EMsg)


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

            Case "Load Spreadsheet"
                Load_Spreadsheet()

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                dst.Tables("RSTNMTMP").Rows.Clear()
                dst.Tables("RSTITEMX").AcceptChanges()

                For Each rowEDT852TX As DataRow In dst.Tables("EDT852TX").Rows
                    rowEDT852TX.Item("SELECTED") = "1"
                Next

                Update_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Load Spreadsheet").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdRSTNMTMP.Visible = ScreenMode
        SplitContainer1.Visible = ScreenMode

        'Set_Read_Only(grpARTCUST1, False)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("RSTNMTMP").Rows.Clear()
        dst.Tables("RSTITEMX").Rows.Clear()
        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("RSTRETL1").Rows.Clear()
        dst.Tables("EDT852TX").Rows.Clear()
        dst.Tables("ARTCUST2").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.cmbFor("OPS_YYYYWW").Text = ""

        BAD_ITEMS_CHECKED = False

    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        OPS_YYYYWW = HFs("OPS_YYYYWW") ' Absx1.cmbFor("HFs("OPS_YYYYWW")").Value & ""
        CUST_CODE = HFs("CUST_CODE")

        rowGLTPARM3 = LookUp("GLTPARM3", OPS_YYYYWW)
        OPS_YYYYPP = rowGLTPARM3.Item("YYYYPP")
        WEEK_END_DATE = rowGLTPARM3.Item("WEEK_END_DATE")

        EnforceConstraints(False)

        Fill_Records("RSTITEMX", CUST_CODE)
        Fill_Records("ARTCUST2", CUST_CODE)

        ASCDATA1.ExecuteSQL("Truncate Table " & EDT852TX)
        ASCDATA1.ExecuteSQL("Insert into " & EDT852TX _
        & " (EDI_DOC_SEQ_NO,EDI_DEPT_NO,OPS_YYYYWW,CUST_CODE,COLLECTION_CODE) " _
        & sqlEDT852TX, "VV", New Object() {CUST_CODE, OPS_YYYYWW})

        ASCDATA1.ExecuteSQL( _
        "Begin Declare Cursor C1 is " _
        & "Select EDT852TX.EDI_DOC_SEQ_NO" _
        & ", SUM (RSTRETL1.QTY_SOLD) QTY_SOLD" _
        & ", SUM (RSTRETL1.AMT_SOLD) AMT_SOLD" _
        & ", SUM (RSTRETL1.QTY_EOW) QTY_EOW" _
        & " from " & EDT852TX & " EDT852TX,RSTRETL1" _
        & " where RSTRETL1.EDI_DOC_SEQ_NO = EDT852TX.EDI_DOC_SEQ_NO " _
        & " group by EDT852TX.EDI_DOC_SEQ_NO;" _
        & " Begin For R1 in C1 Loop " _
        & "  Update " & EDT852TX _
        & "   Set QTY_SOLD = R1.QTY_SOLD, AMT_SOLD = R1.AMT_SOLD, QTY_EOW = R1.QTY_EOW " _
        & "   where EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO; " _
        & " End Loop; " _
        & " End; End;")

        Fill_Records("EDT852TX")

        EnforceConstraints(True)

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Bad Items Only"), UltraWinToolbars.StateButtonTool)
        tlb_sbt.Checked = False
        grdRSTNMTMP.Tag = ""

        Dim dvw As DataView = DirectCast(grdRSTNMTMP.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        For Each rowRSTNMTMP As DataRow In dst.Tables("RSTNMTMP").Select("", "", DataViewRowState.ModifiedCurrent)
            Dim DEPT As String = rowRSTNMTMP.Item("DEPT")
            Dim CLS As String = rowRSTNMTMP.Item("CLS")
            Dim STYLE As String = rowRSTNMTMP.Item("STYLE")
            Dim ITEM_CODE As String = rowRSTNMTMP.Item("ITEM_CODE")
            Dim rowRSTITEMX As DataRow = dst.Tables("RSTITEMX").Rows.Find _
            (New String() {CUST_CODE, DEPT & CLS & STYLE})
            If rowRSTITEMX Is Nothing Then
                rowRSTITEMX = dst.Tables("RSTITEMX").NewRow
                rowRSTITEMX.Item("CUST_CODE") = CUST_CODE
                rowRSTITEMX.Item("CUST_ITEM_CODE") = DEPT & CLS & STYLE
                dst.Tables("RSTITEMX").Rows.Add(rowRSTITEMX)
            End If
            rowRSTITEMX.Item("ITEM_CODE") = ITEM_CODE
        Next

        Update_Record_TDA("RSTITEMX")

        Update_Record_TDA("RSTNMTMP", "CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYWW = '" & OPS_YYYYWW & "'")


        Dim EDI_DOC_SEQ_NO As String = ""

        ' Delete All Manually Entered Retail Sales Documents for this Customer and Week
        For Each rowEDT852TX As DataRow In dst.Tables("EDT852TX").Select("SELECTED = '1'")
            EDI_DOC_SEQ_NO = rowEDT852TX.Item("EDI_DOC_SEQ_NO")
            TAC.RSCMAIN1.Update_RSTRETLx(EDI_DOC_SEQ_NO, "-")
            ASCDATA1.ExecuteSQL("Delete from RSTRETL1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            ASCDATA1.ExecuteSQL("Delete from EDT852T1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
        Next


        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("RSTRETL1").Rows.Clear()

        EDI_DOC_SEQ_NO = ""
        Dim DEPT_CODE As String = ""
        Dim COLLECTION_CODE As String = ""

        For Each row As DataRow In dst.Tables("RSTNMTMP").Select("", "DEPT,COLLECTION_CODE")
            If EDI_DOC_SEQ_NO = "" Or DEPT_CODE <> row.Item("DEPT") & "" Then
                'COLLECTION_CODE = row.Item("COLLECTION_CODE")
                DEPT_CODE = row.Item("DEPT")
                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO")
                Dim rowEDT852T1 As DataRow = dst.Tables("EDT852T1").NewRow
                rowEDT852T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                rowEDT852T1.Item("EDI_FROM_DATE") = WEEK_END_DATE.AddDays(-6)
                rowEDT852T1.Item("EDI_TO_DATE") = WEEK_END_DATE
                rowEDT852T1.Item("EDI_STATUS") = "M"
                rowEDT852T1.Item("EDI_DEPT_NO") = DEPT_CODE
                rowEDT852T1.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowEDT852T1.Item("OPS_YYYYWW") = OPS_YYYYWW
                rowEDT852T1.Item("CUST_CODE") = CUST_CODE
                rowEDT852T1.Item("INIT_DATE") = DATETIME_STAMP
                rowEDT852T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowEDT852T1.Item("BRAND_CODE") = "JH"
                rowEDT852T1.Item("DATA_LEVEL") = "I"
                dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)

            End If

            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")

            Dim rowRSTRETL1 As DataRow
            rowRSTRETL1 = dst.Tables("RSTRETL1").Rows.Find _
            (New String() {EDI_DOC_SEQ_NO, CUST_CODE, CUST_STORE_NO, ITEM_CODE})

            If rowRSTRETL1 Is Nothing Then
                rowRSTRETL1 = dst.Tables("RSTRETL1").NewRow
                With rowRSTRETL1
                    .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("ITEM_CODE") = ITEM_CODE

                    .Item("QTY_SOLD") = row.Item("QTY_SOLD")
                    .Item("AMT_SOLD") = row.Item("AMT_SOLD")
                    .Item("QTY_EOW") = row.Item("QTY_OH")
                    .Item("OPS_YYYYPP") = OPS_YYYYPP
                    .Item("OPS_YYYYWW") = OPS_YYYYWW
                    'rowRSTRETL1.Item("COLLECTION_CODE") = COLLECTION_CODE
                    dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                End With
            Else
                With rowRSTRETL1
                    .Item("QTY_SOLD") = Val(.Item("QTY_SOLD") & "") + Val(row.Item("QTY_SOLD") & "")
                    .Item("AMT_SOLD") = Val(.Item("AMT_SOLD") & "") + Val(row.Item("AMT_SOLD") & "")
                    .Item("QTY_EOW") = Val(.Item("QTY_EOW") & "") + Val(row.Item("QTY_OH") & "")
                End With
            End If
        Next

        If EDI_DOC_SEQ_NO <> "" Then
            Update_Record_TDA("EDT852T1")
            Update_Record_TDA("RSTRETL1")
            TAC.RSCMAIN1.Update_RSTRETLx(EDI_DOC_SEQ_NO, "+")
        End If

        Call CommitTrans("Update Complete")

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
        Call Load_Popup_Menu(grdRSTNMTMP, "SSSBB", "Show Filter", "Show GroupBox", "Show Bad Items Only", "Delete Bad Items Only", "Delete All Records")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If tlb_pop.Tools.Exists("Show Bad Items Only") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Bad Items Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.Tag <> "")
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing OrElse Not GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Show Bad Items Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim dvw As DataView = DirectCast(grd.DataSource, DataTable).DefaultView
                If tlb_sbt.Checked Then
                    If Not BAD_ITEMS_CHECKED Then Check_Items()
                    grd.Tag = "ISNULL(BAD_ITEM,'X') <> 'X'"
                Else
                    grd.Tag = ""
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Setting Filter")
                dvw.RowFilter = grd.Tag
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Delete Bad Items Only"
                Check_Items()
                ASCDATA1.DeleteRows("RSTNMTMP", "ISNULL(BAD_ITEM,'X') <> 'X'")

            Case "Delete All Records"
                dst.Tables("RSTNMTMP").Rows.Clear()

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

    'Function Load_Excel_Sheetx() As Boolean

    '    'uses Gembox - requires a reference to Gembox.Spreadsheet

    '    Load_Excel_Sheetx = False

    '    Try

    '        Dim xlBook As GemBox.Spreadsheet.ExcelFile = New GemBox.Spreadsheet.ExcelFile
    '        Dim xlSheet As GemBox.Spreadsheet.ExcelWorksheet
    '        'FILENAME = "C:\GemBox\GemBox.Spreadsheet Free 3.1\Samples\NET20\VB\ReadingSamplesVB\TestWorkbook.xls"
    '        'FILENAME = "TestWorkbook.xls"
    '        xlBook.LoadXls(FILENAME)

    '        xlSheet = xlBook.Worksheets(0)
    '        If xlBook.Worksheets.Count > 1 Then
    '            Dim f As New ASFMSGBF
    '            Dim tbl As New DataTable
    '            tbl.Columns.Add("WorkSheet")
    '            For Each xlSheet In xlBook.Worksheets
    '                tbl.Rows.Add(xlSheet.Name)
    '            Next
    '            f.Show_grd(tbl, Me, "Worksheets in Workbook")
    '            ASCMAIN1.response = f.user_option
    '            f.Dispose()
    '        End If

    '        'For Each xls_Sheet In xls_File.Worksheets
    '        '    Console.WriteLine("--------- {0} ---------", sheet.Name)
    '        '    Dim i As Int32 = 0
    '        '    For Each row In sheet.Rows
    '        '        For Each cell In row.AllocatedCells
    '        '            If Not cell.Value Is Nothing Then
    '        '                Console.Write("{0}({1})", cell.Value, cell.Value.GetType().Name)
    '        '            End If
    '        '            Console.Write(vbTab)
    '        '        Next
    '        '    Next
    '        'Next

    '        dst.Tables("RSTNMTMP").Rows.Clear()

    '        'Dim DEPT As String = ""
    '        Dim STORE As String = ""
    '        Dim WE_DATE As Date = Nothing

    '        Dim r As Int32 = 0

    '        Do While xlSheet.Cells(r, 1).Value & "" <> "" Or xlSheet.Cells(r + 1, 1).Value & "" <> "" Or xlSheet.Cells(r + 2, 1).Value & "" <> ""
    '            If xlSheet.Cells(r, 5).Value & "" = "Vendor:" Then
    '                r += 1
    '            End If
    '            If xlSheet.Cells(r, 5).Value & "" = "Vendor Name:" Then
    '                r += 2
    '                STORE = xlSheet.Cells(r, 6).Value
    '                WE_DATE = Split(xlSheet.Cells(r, 8).Value, " ")(0)

    '                If Format(WE_DATE, "ddMMyyyy") <> Format(WEEK_END_DATE, "ddMMyyyy") Then
    '                    MsgBox("W/E Date in Spreadsheet (" & Format(WE_DATE, "MM/dd/yyyy") & ") does not match W/E Date for Week Selected (" & Format(WEEK_END_DATE, "MM/dd/yyyy") & ")", MsgBoxStyle.OkOnly, "Spreadsheet Load will be Terminated")
    '                    Load_Excel_Sheetx = False
    '                    Return False
    '                End If
    '                r += 5
    '            Else
    '                If xlSheet.Cells(r, 1).Value & "" <> "" Then
    '                    If STORE = "" Or xlSheet.Cells(r, 8).Value & "" = "Total" Then Stop
    '                    If Len(xlSheet.Cells(r, 6).Value & "") > 2 Then Stop
    '                    Dim rowRSTNMTMP As DataRow = dst.Tables("RSTNMTMP").NewRow
    '                    rowRSTNMTMP.Item("DEPT") = xlSheet.Cells(r, 5).Value
    '                    rowRSTNMTMP.Item("STORE") = STORE
    '                    rowRSTNMTMP.Item("WE_DATE") = WE_DATE
    '                    rowRSTNMTMP.Item("XNO") = XNO
    '                    rowRSTNMTMP.Item("CLS") = xlSheet.Cells(r, 6).Value
    '                    rowRSTNMTMP.Item("STYLE") = xlSheet.Cells(r, 8).Value
    '                    rowRSTNMTMP.Item("STYLE_DESC") = xlSheet.Cells(r, 9).Value
    '                    rowRSTNMTMP.Item("QTY_SOLD") = Val(xlSheet.Cells(r, 12).Value & "")
    '                    rowRSTNMTMP.Item("QTY_OH") = Val(xlSheet.Cells(r, 16).Value & "")
    '                    dst.Tables("RSTNMTMP").Rows.Add(rowRSTNMTMP)
    '                End If
    '            End If
    '            r += 1
    '            If r Mod 10 = 0 Then
    '                ASCMAIN1.Progress("Row " & CStr(r))
    '            End If
    '        Loop
    '        'Stop

    '    Catch ex As Exception

    '        MsgBox("Error reading spreadsheet" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Cannot Process Spreadsheet")
    '    End Try
    '    Load_Excel_Sheetx = True

    'End Function

    Function Load_Excel_Sheet() As Boolean

        Application.DoEvents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Excel Sheet")

        Load_Excel_Sheet = False

        'MsgBox("Warning - currently picking up PTD not WTD")
        '' for QTY_SOLD, should be using row.Item(12) not row.Item(13) (see code below)
        '' - 13 is temp to pick up PTD to load catch-up spreadsheets

        Try

            Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                "data source=" & FILENAME & ";" & _
                "Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""

            Using objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                objConnection.Open()

                Dim strSQL As String = "SELECT * FROM [" & EXCEL_SHEET & "]"
                'Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                Dim dt As New DataTable
                objAdapter.Fill(dt)
                objConnection.Close()

                dst.Tables("RSTNMTMP").Rows.Clear()

                Dim STORE As String = ""
                Dim WE_DATE As Date = Nothing

                For R As Int32 = 0 To dt.Rows.Count - 1
                    Dim row As DataRow = dt.Rows(R)
                    If row.Item(5) & "" = "Vendor Name:" Then
                        R = R + 2
                        row = dt.Rows(R)
                        STORE = row.Item(6)
                        R = R + 1
                        row = dt.Rows(R)
                        WE_DATE = Split(row.Item(8), " ")(0)
                        R = R + 3
                    Else
                        If row.Item(0) & "" <> "" Then
                            Dim rowRSTNMTMP As DataRow = dst.Tables("RSTNMTMP").NewRow
                            rowRSTNMTMP.Item("DEPT") = row.Item(5)
                            rowRSTNMTMP.Item("STORE") = STORE
                            rowRSTNMTMP.Item("WE_DATE") = WE_DATE
                            rowRSTNMTMP.Item("XNO") = XNO
                            rowRSTNMTMP.Item("CLS") = row.Item(6)
                            rowRSTNMTMP.Item("STYLE") = row.Item(8)
                            rowRSTNMTMP.Item("STYLE_DESC") = row.Item(9)
                            rowRSTNMTMP.Item("QTY_SOLD") = Val(Replace(Replace(row.Item(12) & "", "(", "-"), ")", "")) ' WTD
                            'rowRSTNMTMP.Item("QTY_SOLD") = Val(Replace(Replace(row.Item(13) & "", "(", "-"), ")", "")) ' PTD
                            rowRSTNMTMP.Item("QTY_OH") = Val(Replace(Replace(row.Item(16) & "", "(", "-"), ")", ""))
                            dst.Tables("RSTNMTMP").Rows.Add(rowRSTNMTMP)
                        End If
                    End If

                    If R Mod 10 = 0 Then
                        ASCMAIN1.Progress("-", CStr(R))
                    End If

                Next

            End Using

            Load_Excel_Sheet = True

        Catch ex As Exception

            MsgBox("Error reading spreadsheet" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Cannot Process Spreadsheet")
        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Function

    'Function Load_Excel_Sheeto() As Boolean
    '    ' uses InterOp and requires a reference to Microsoft.Office.Interop.Excel

    '    Load_Excel_Sheeto = False
    '    'Stop ' USE CELLS 13 NOT 14 - 14 IS FOR TEMP SEASONAL
    '    Try

    '        Dim xlExcel As New Microsoft.Office.Interop.Excel.Application
    '        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    '        'FILENAME = "C:\TestWorkbook.xls"
    '        xlBook = GetObject(FILENAME)
    '        'xlBook.Application.Visible = True
    '        'xlBook.Windows(1).Visible = True
    '        'xlBook.Application.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized

    '        'xlExcel = xlExcel.Workbooks.Open(FILENAME)
    '        'xlExcel = xlExcel.Worksheets(1)
    '        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet = xlBook.Worksheets(1)

    '        'Dim strSQL As String = "SELECT * FROM [" & EXCEL_SHEET & "]"
    '        'Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
    '        'Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
    '        'Dim dt As New DataTable
    '        'objAdapter.Fill(dt)
    '        'objConnection.Close()

    '        dst.Tables("RSTNMTMP").Rows.Clear()

    '        'Dim DEPT As String = ""
    '        Dim STORE As String = ""
    '        Dim WE_DATE As Date = Nothing

    '        Dim r As Int32 = 0
    '        r = 1

    '        Do While xlSheet.Cells(r, 2).Value & "" <> "" Or xlSheet.Cells(r + 1, 2).Value & "" <> "" Or xlSheet.Cells(r + 2, 2).Value & "" <> ""
    '            If xlSheet.Cells(r, 6).Value & "" = "Vendor:" Then
    '                r += 1
    '            End If
    '            If xlSheet.Cells(r, 6).Value & "" = "Vendor Name:" Then
    '                r += 2
    '                STORE = xlSheet.Cells(r, 7).Value
    '                WE_DATE = Split(xlSheet.Cells(r, 9).Value, " ")(0)

    '                If Format(WE_DATE, "ddMMyyyy") <> Format(WEEK_END_DATE, "ddMMyyyy") Then
    '                    MsgBox("W/E Date in Spreadsheet (" & Format(WE_DATE, "MM/dd/yyyy") & ") does not match W/E Date for Week Selected (" & Format(WEEK_END_DATE, "MM/dd/yyyy") & ")", MsgBoxStyle.OkOnly, "Spreadsheet Load will be Terminated")
    '                    Load_Excel_Sheeto = False
    '                    Return False
    '                End If
    '                r += 5
    '            Else
    '                If xlSheet.Cells(r, 2).Value & "" <> "" Then
    '                    If STORE = "" Or xlSheet.Cells(r, 9).Value & "" = "Total" Then Stop
    '                    If Len(xlSheet.Cells(r, 7).Value & "") > 2 Then Stop
    '                    Dim rowRSTNMTMP As DataRow = dst.Tables("RSTNMTMP").NewRow
    '                    rowRSTNMTMP.Item("DEPT") = xlSheet.Cells(r, 6).Value
    '                    rowRSTNMTMP.Item("STORE") = STORE
    '                    rowRSTNMTMP.Item("WE_DATE") = WE_DATE
    '                    rowRSTNMTMP.Item("XNO") = XNO
    '                    rowRSTNMTMP.Item("CLS") = xlSheet.Cells(r, 7).Value
    '                    rowRSTNMTMP.Item("STYLE") = xlSheet.Cells(r, 9).Value
    '                    rowRSTNMTMP.Item("STYLE_DESC") = xlSheet.Cells(r, 10).Value
    '                    rowRSTNMTMP.Item("QTY_SOLD") = Val(xlSheet.Cells(r, 14).Value & "")
    '                    rowRSTNMTMP.Item("QTY_OH") = Val(xlSheet.Cells(r, 17).Value & "")
    '                    dst.Tables("RSTNMTMP").Rows.Add(rowRSTNMTMP)
    '                End If
    '            End If
    '            r += 1
    '            If r Mod 10 = 0 Then
    '                ASCMAIN1.Progress("Row " & CStr(r))
    '            End If
    '            ' If r > 150 Then Exit Do
    '        Loop

    '        xlSheet = Nothing
    '        xlBook = Nothing
    '        xlExcel = Nothing

    '    Catch ex As Exception

    '        MsgBox("Error reading spreadsheet" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Cannot Process Spreadsheet")
    '    End Try
    '    Load_Excel_Sheeto = True

    'End Function

    Private Sub grdRSTNMTMP_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTNMTMP.ClickCellButton
        Dim sql_where As String = ""
        If grdRSTNMTMP.ActiveCell Is Nothing Then
            Exit Sub
        End If
        Select Case grdRSTNMTMP.ActiveCell.Column.Key
            Case "CUST_STORE_NO"
                grdClickCellButton(grdRSTNMTMP, "CUST_CODE = '" & CUST_CODE & "'")
            Case "ITEM_CODE"
                grdClickCellButton(grdRSTNMTMP)
        End Select
    End Sub

    Sub Check_Items(Optional ByRef EMsg As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Checking Items")

        dst.Tables("RSTNMTMP").AcceptChanges()
        For Each rowRSTNMTMP As DataRow In dst.Tables("RSTNMTMP").Rows
            rowRSTNMTMP.Item("BAD_ITEM") = DBNull.Value
        Next

        Dim COLLECTION_CODEs_bad As String = ""
        Dim ITEM_CODEs_bad As String = ""
        For Each row As DataRow In ASCDATA1.SelectDistinct("RSTNMTMP", "ITEM_CODE").Rows
            Dim BAD_ITEM As String = ""
            Dim ITEM_CODE As String = row.Item(0) & ""
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 Is Nothing Then
                ITEM_CODEs_bad &= "," & IIf(ITEM_CODE = "", "{Blank}", ITEM_CODE)
                BAD_ITEM = "I"
            End If

            If ITEM_CODE = "" Then
                For Each rowRSTNMTMP As DataRow In dst.Tables("RSTNMTMP").Select("ITEM_CODE IS NULL")
                    rowRSTNMTMP.Item("BAD_ITEM") = BAD_ITEM
                Next
            Else
                For Each rowRSTNMTMP As DataRow In dst.Tables("RSTNMTMP").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    If BAD_ITEM = "I" Then
                        rowRSTNMTMP.Item("BAD_ITEM") = BAD_ITEM
                    Else
                        rowRSTNMTMP.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE") & ""
                        rowRSTNMTMP.Item("ITEM_RETAIL_PRICE") = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                        If rowRSTNMTMP.Item("COLLECTION_CODE") & "" = "" Then
                            COLLECTION_CODEs_bad &= "," & IIf(ITEM_CODE = "", "{Blank}", ITEM_CODE)
                            rowRSTNMTMP.Item("BAD_ITEM") = "C"
                        End If
                    End If
                Next
            End If
        Next

        If ITEM_CODEs_bad <> "" Then
            EMsg &= vbCr & "Bad Items: " & Mid(ITEM_CODEs_bad, 2)
        End If
        If COLLECTION_CODEs_bad <> "" Then
            EMsg &= vbCr & "Items with Bad Collection Code: " & Mid(COLLECTION_CODEs_bad, 2)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        BAD_ITEMS_CHECKED = True
    End Sub

    Sub Load_Spreadsheet()

        EnforceConstraints(False)

        If Not Load_Excel_Sheet() Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Processing Excel Data")

        Dim OPS_YYYYWW As String = Absx1.cmbFor("OPS_YYYYWW").Value
        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", OPS_YYYYWW)

        For Each rowRSTNMTMP As DataRow In dst.Tables("RSTNMTMP").Select("", "STORE", DataViewRowState.CurrentRows)
            Dim STORE As String = rowRSTNMTMP.Item("STORE") & ""
            rowRSTNMTMP.Item("CUST_STORE_NO") = "000" & Mid(STORE, 1, 3)
            Dim DEPT As String = rowRSTNMTMP.Item("DEPT") & ""
            Dim CLS As String = rowRSTNMTMP.Item("CLS") & ""
            Dim STYLE As String = rowRSTNMTMP.Item("STYLE") & ""
            Dim rowRSTITEMX As DataRow = dst.Tables("RSTITEMX").Rows.Find(New String() {CUST_CODE, DEPT & CLS & STYLE})
            If rowRSTITEMX IsNot Nothing Then
                rowRSTNMTMP.Item("ITEM_CODE") = rowRSTITEMX.Item("ITEM_CODE")
            End If
            rowRSTNMTMP.Item("XNO") = XNO
            rowRSTNMTMP.Item("CUST_CODE") = CUST_CODE
            rowRSTNMTMP.Item("OPS_YYYYWW") = OPS_YYYYWW
        Next
        dst.Tables("RSTNMTMP").AcceptChanges()

        EnforceConstraints(True)

        ASCDATA1.DeleteRows("RSTNMTMP", "ISNULL(QTY_SOLD,0) = 0 AND ISNULL(QTY_OH,0) = 0")

        Dim EMsg As String = ""
        For Each row As DataRow In ASCDATA1.SelectDistinct("RSTNMTMP", "WE_DATE").Rows
            Dim WE_DATE As Date = row.Item(0)
            If Format(WE_DATE, "yyyyMMdd") <> Format(rowGLTPARM3.Item("WEEK_END_DATE"), "yyyyMMdd") Then
                EMsg &= vbCr & "W/E date " & Format(WE_DATE, "MM/dd/yyyy") & " in grid does not agree with W/E date for Week " & OPS_YYYYWW & " (" & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yyyy") & ")"
            End If
        Next
        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Warning: Update will Not be permitted")
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")

    End Sub
End Class