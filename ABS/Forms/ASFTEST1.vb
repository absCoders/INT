Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json
Imports nsoftware.IPWorksSSH
'Imports System.Security.Cryptography
Public Class ASFTEST1
    Dim dst As DataSet
    Dim theLog As String = ""

    Private Sub ASFTEST1_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click

        ASCMAIN1.sql = "Select * from ASTSECM1"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        dst.Tables.Add(tbl)

        Dim RPT_FILENAME As String = ASCMAIN1.Folders("Reports") & "ASRTEST1.RPT"
        If ASCMAIN1.Running_in_VS Then
            Dim XSD_FILENAME As String = ASCMAIN1.Folders("Temp") & "ASFTEST1.XSD"
            If Not My.Computer.FileSystem.FileExists(XSD_FILENAME) Then
                dst.WriteXml(XSD_FILENAME, XmlWriteMode.WriteSchema)
            End If
        End If

        ASCMAIN1.CR_RPT.Load(RPT_FILENAME)

        ASCMAIN1.CR_RPT.SetDataSource(dst)

        For Each sr As CrystalDecisions.CrystalReports.Engine.ReportDocument In ASCMAIN1.CR_RPT.Subreports
            Try
                sr.SetDataSource(dst)
            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If
            End Try
        Next

        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")
        Dim FILENAME As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".RPT"
        Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions
        DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & FILENAME

        With ASCMAIN1.CR_RPT.ExportOptions
            .DestinationOptions = DestOpt
            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport

            'Select Case ExportFormat
            '    Case "RPT"
            '        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
            '    Case "PDF"
            '        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
            'End Select
        End With













        Dim crv As New CrystalDecisions.Windows.Forms.CrystalReportViewer
        ' Add a Crystal Report Viewer to the Tab Page Control & Configure it
        SplitContainer1.Panel2.Controls.Add(crv)
        crv.ActiveViewIndex = -1
        crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        crv.Dock = System.Windows.Forms.DockStyle.Fill
        crv.BackColor = System.Drawing.Color.FromArgb(222, 223, 206)

        Dim REPORT_FILENAME As String = ASCMAIN1.Folders("Temp") & FILENAME

        Dim RPT As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Try
            RPT.Load(REPORT_FILENAME)
            crv.ReportSource = RPT

        Catch ex As Exception
            MsgBox("Problem Report: " & ASCMAIN1.CR_RPT.FileName & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load Report " & REPORT_NO)

        End Try

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click

        Dim F As ASFBASE1 = ASCMAIN1.ActiveForm
        F.remotely_controlled = True
        If Not F.ScreenMode Then
            F.Click_Command("Cancel")
        End If

        ASCMAIN1.sql = "SELECT * FROM MM_TAX_ADJ"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            F.Absx1.txtFor("CUST_CODE").Text = row.Item("CUST")
            F.Absx1.txtFor("CUST_SHIP_TO_NO").Text = row.Item("SHIP_TO") & ""
            Dim TAX_ADJ As Decimal = Val(row.Item("TAX_ADJ"))
            F.Absx1.optFor("INV_TYPE").CheckedIndex = IIf(TAX_ADJ < 0, 0, 1)
            F.Click_Command("New")
            If F.ScreenMode Then
                F.Absx1.numFor("INV_STAX").Value = TAX_ADJ
                F.Absx1.txtFor("STAX_CODE").Text = row.Item("STAX_CD")
                F.Absx1.txtFor("INV_NOTES").Text = "Adj tax doc#" & row.Item("INV_NO")
                F.Absx1.txtFor("ORDR_CUST_PO").Text = row.Item("INV_NO")
                'Dim X As MsgBoxResult = MsgBox("UPDATE", MsgBoxStyle.YesNoCancel, "VERIFICATION")
                'If X = MsgBoxResult.Cancel Then
                '    Exit Sub
                'ElseIf X = MsgBoxResult.Yes Then
                '    F.Click_Command("Update")
                'End If
                F.Click_Command("Update")
            Else
                MsgBox("X", MsgBoxStyle.OkOnly, "X")
            End If
        Next
    End Sub

    Private Sub btnXLS_Click(sender As Object, e As EventArgs) Handles btnXLS.Click

        'Dim FILENAME2 = "C:\ABS\DEL.XLSX"
        'Dim oWB2 As SpreadsheetGear.IWorkbook
        'oWB2 = SpreadsheetGear.Factory.GetWorkbook(FILENAME2)
        'Dim oSheet2 As SpreadsheetGear.IWorksheet = oWB2.Worksheets(2)
        'Dim condition2 As SpreadsheetGear.IFormatCondition = oSheet2.Cells("G6").FormatConditions(0)


        Dim FILENAME = ASCMAIN1.Next_Control_No("ASFTEST1") & ".XLSX"
        Dim TY As String = "2016"
        Dim LY As String = "2015"

        Dim oWB As SpreadsheetGear.IWorkbook
        oWB = SpreadsheetGear.Factory.GetWorkbook()
        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i
        Dim oSheet As SpreadsheetGear.IWorksheet
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        Dim SI As Integer = 0

        Dim maxCols As Integer = 10

        Dim Row0 As Integer = 0
        Dim Col0 As Integer = 1

        Dim Heading_Row0 As Integer
        Dim Data_Row_Count As Integer
        Dim Data_Col_Count As Integer
        Dim Totals_Row0 As Integer

        oSheet = oWB.Sheets(SI)
        oSheet.Name = "Summary EWS"
        oSheet.Tab.Color = SpreadsheetGear.Color.FromArgb(146, 208, 80)  ' SpreadsheetGear.Colors.LightGreen

        oSheet.Cells(Row0 + 0, 0).EntireColumn.ColumnWidth = 1

        With oSheet.Cells(Row0 + 0, Col0)
            .Value = "ABB OPTICAL GROUP"
            '  .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            .Font.Name = "Times New Roman"
            .Font.Size = 18
            .GetCharacters(0, 3).Font.Bold = True
        End With

        With oSheet.Cells(Row0 + 1, Col0, 1, Col0 + maxCols - 1)
            '.Merge()
            '.VerticalAlignment = SpreadsheetGear.VAlign.Center
            .Borders.Weight = SpreadsheetGear.BorderWeight.Thick
            '.Borders.Color = SpreadsheetGear.Colors.Blue
            .Borders.Color = SpreadsheetGear.Color.FromArgb(0, 112, 192)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.None
            .EntireRow.RowHeight = 3.5
        End With

        With oSheet.Cells(Row0 + 2, Col0)
            .Value = "DEL SALES BY EYEWEAR SPECIALIST"
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = True
        End With

        With oSheet.Cells(Row0 + 2, Col0 + maxCols - 1)
            .Value = "As of " & Format(Now, "D")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With


        ' a report starts here


        ' Notes on the DataTable
        ' 1) Arrange Columns of DataTable in the sequence desired in the XLS
        ' 2) Use placeholders for formula fields, setting expressions after the datatable has been created
        ' 3) use all Upper Case, and avoid the use of spaces in expressions
        ' 4) avoid column names that look like R1C1 syntax

        Dim tbl As New DataTable
        With tbl.Columns
            .Add("SREP2_CODE", GetType(System.String))
            tbl.Columns("SREP2_CODE").MaxLength = 6
            .Add("SREP2_NAME", GetType(System.String))
            tbl.Columns("SREP2_NAME").MaxLength = 30
            .Add("TY_YTD_SALES", GetType(System.Decimal))
            .Add("LY_YTD_SALES", GetType(System.Decimal))
            .Add("YTD_SALES_GROWTH", GetType(System.Decimal), "ISNULL(TY_YTD_SALES,0)-ISNULL(LY_YTD_SALES,0)")
            .Add("YTD_SALES_GROWTH_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_YTD_SALES,0)=0,0,100*YTD_SALES_GROWTH/LY_YTD_SALES)")
        End With

        With tbl
            .Columns("SREP2_CODE").Caption = "Rep2 #"
            .Columns("SREP2_NAME").Caption = "Rep2 Name"
            .Columns("TY_YTD_SALES").Caption = "YTD Sales|{TY}"
            .Columns("LY_YTD_SALES").Caption = "YTD Sales|{LY}"
            .Columns("YTD_SALES_GROWTH").Caption = "YTD Sales|Growth $"
            .Columns("YTD_SALES_GROWTH_PCT").Caption = "YTD Sales|Growth %"
        End With

        Dim tblSchema As New DataTable
        With tblSchema.Columns
            .Add("COLUMN_NAME", GetType(System.String))
            .Add("COLUMN_INDEX", GetType(System.Int32))
            .Add("COLUMN_LENGTH", GetType(System.Int32), "LEN(COLUMN_NAME)")
        End With

        For iCol As Integer = 0 To tbl.Columns.Count - 1
            tblSchema.Rows.Add(New Object() {tbl.Columns(iCol).ColumnName, iCol})
        Next

        tbl.Rows.Add(New Object() {"003", "SUSAN MACKINNON", 834994, 719691})
        tbl.Rows.Add(New Object() {"006", "LAURIE MOGCK", 1684939, 1192238})
        tbl.Rows.Add(New Object() {"007", "SCOTT ROSENWALD", 3523908, 3326836})
        tbl.Rows.Add(New Object() {"011", "DAN LASNER", 3601640, 1853768})
        tbl.Rows.Add(New Object() {"012", "FLYNN GILDEN", 2433891, 1976070})
        tbl.Rows.Add(New Object() {"020", "DAVID SINGER", 1189255, 516350})
        tbl.Rows.Add(New Object() {"022", "NICK FOLKERS", 3150641, 2143524})
        tbl.Rows.Add(New Object() {"024", "BART VANDER VELDE", 1234780, 1037199})
        tbl.Rows.Add(New Object() {"025", "MATTHEW RUPPERT", 2192096, 2858721})
        tbl.Rows.Add(New Object() {"026", "LORI McKENZIE", 769279, 814540})
        tbl.Rows.Add(New Object() {"032", "0", 1063498, 1735302})



        Heading_Row0 = Row0 + 4
        Data_Row_Count = tbl.Rows.Count
        Data_Col_Count = tbl.Columns.Count
        Totals_Row0 = Heading_Row0 + Data_Row_Count + 1 + 1

        oSheet.Range(Excel_Cell0(Heading_Row0, Col0)).EntireRow.RowHeight = 43.5
        With oSheet.Cells(Heading_Row0, Col0, Heading_Row0, Col0 + Data_Col_Count - 1)
            .Borders.Weight = SpreadsheetGear.BorderWeight.Thin
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            '  .Borders.Color = SpreadsheetGear.Colors.Blue
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Font.Bold = True
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.None
            .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.None
        End With

        range = oSheet.Cells(Heading_Row0 + 1, Col0, Heading_Row0 + Data_Row_Count + 2, Col0)
        range.EntireRow.RowHeight = 14.5

        For iCol As Integer = 0 To tbl.Columns.Count - 1
            Dim dcol As DataColumn = tbl.Columns(iCol)

            Dim C As String = dcol.Caption
            C = Replace(C, "{TY}", TY)
            C = Replace(C, "{LY}", LY)
            C = Replace(C, "|", vbCrLf)

            With oSheet.Cells(Heading_Row0, Col0 + iCol)
                .Value = C
                If C.Contains("%") Then
                    .Interior.Color = SpreadsheetGear.Color.FromArgb(242, 242, 242)
                    oSheet.Cells(Totals_Row0, Col0 + iCol).Interior.Color = SpreadsheetGear.Color.FromArgb(242, 242, 242)
                End If
            End With

            If dcol.ColumnName = "SREP2_CODE" Then
                oSheet.Cells(Heading_Row0 + 1, Col0 + iCol, Heading_Row0 + Data_Row_Count, Col0 + iCol).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If

            range = oSheet.Cells(Heading_Row0 + 1, Col0 + iCol, Heading_Row0 + Data_Row_Count + 2, Col0 + iCol)

            If dcol.DataType.ToString = "System.Decimal" Then

                Dim XC1 As String = Excel_Cell0(Heading_Row0 + 1, Col0 + iCol)
                Dim XC2 As String = Excel_Cell0(Heading_Row0 + Data_Row_Count, Col0 + iCol)

                Dim F As String = dcol.Expression
                If F <> "" Then
                    F = Replace(F, "IIF(", "IF(")
                    Do While F.Contains("ISNULL(")
                        Dim x As Integer = InStr(F, "ISNULL(")
                        Dim y As Integer = InStr(Mid(F, x), ",0)")
                        F = Mid(F, 1, x - 1) & Mid(F, x + 6 + 1, y - (6 + 1) - 1) & Mid(F, x + y + 3 - 1)
                    Loop

                    For Each row As DataRow In tblSchema.Select("", "COLUMN_LENGTH DESC")
                        Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
                        Dim COLUMN_INDEX As String = Val(row.Item("COLUMN_INDEX") & "")
                        If F.Contains(COLUMN_NAME) Then
                            F = Replace(F, COLUMN_NAME, Excel_Cell0(Totals_Row0, Col0 + COLUMN_INDEX))
                        End If
                    Next

                    oSheet.Cells(Totals_Row0, Col0 + iCol).Formula = "=" & F
                    'rangeCopyFrom = oSheet.Range(Totals_Row0, Col0 + iCol)
                    'rangePaste_To = oSheet.Range(Heading_Row0 + 1, Col0 + iCol, Heading_Row0 + Data_Col_Count, Col0 + iCol)
                    'rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)

                Else
                    oSheet.Cells(Totals_Row0, Col0 + iCol).Formula = String.Format("=SUBTOTAL(9,{0}:{1})", XC1, XC2)
                End If

                If dcol.Caption.Contains("%") Then
                    'range.NumberFormat = "%"
                    range.NumberFormat = "##0.0\%"

                    Dim columnHeadingStyle As SpreadsheetGear.IStyle = oWB.Styles("Heading 2")
                    oSheet.Cells(XC1 & ":" & XC2).Style = columnHeadingStyle
                    Dim condition As SpreadsheetGear.IFormatCondition = oSheet.Cells(XC1 & ":" & XC2).FormatConditions.Add(
             SpreadsheetGear.FormatConditionType.Expression,
             SpreadsheetGear.FormatConditionOperator.Between,
             "=MOD(ROW(),2)=0", Nothing)
                    condition.Interior.Color = columnHeadingStyle.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).Color

                Else
                    range.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ' - '??_);_(@_)" '"$#,##0"
                End If
                range.EntireColumn.ColumnWidth = 13
            Else
                range.NumberFormat = "@"
                range.EntireColumn.ColumnWidth = dcol.MaxLength
            End If
        Next

        range = oSheet.Range(Excel_Cell0(Heading_Row0 + 1, Col0 + 0) & ":" & Excel_Cell0(Heading_Row0 + 1, Col0 + Data_Col_Count - 1))

        Dim dvw As DataView = tbl.DefaultView
        dvw.Sort = "YTD_SALES_GROWTH_PCT DESC"
        Dim tbl2 As DataTable = dvw.ToTable

        range.CopyFromDataTable(tbl2, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)

        With oSheet.Cells(Totals_Row0, Col0 + 0)
            .Value = "TOTAL"
        End With

        For i As Integer = -1 To 0
            With oSheet.Cells(Totals_Row0 + i, Col0 + 0, Totals_Row0 + i, Col0 + Data_Col_Count - 1)
                If i = 0 Then
                    .Borders.Weight = SpreadsheetGear.BorderWeight.Medium
                    .Font.Bold = True
                Else
                    .Borders.Weight = SpreadsheetGear.BorderWeight.Thin
                End If
                '  .Borders.Color = SpreadsheetGear.Colors.Blue
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.None
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.None
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.None
                .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.None
                .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.None
            End With
        Next

        '  range = oSheet.Cells(Heading_Row0 + 1, Col0 + 0, Heading_Row0 + Data_Col_Count + 2, Col0 + Data_Col_Count - 1)
        ' oSheet.Range.AutoFit()
        oSheet.WindowInfo.DisplayGridlines = False
        oSheet.WindowInfo.Zoom = 85
        oWB.WorkbookSet.Calculation = SpreadsheetGear.Calculation.Automatic
        oWB.WorkbookSet.CalculationOnDemand = False
        oWB.WorkbookSet.Calculate()


        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Dim p As Process = Process.Start(FILENAME)
        oWB = Nothing

    End Sub
    ''' <summary>
    ''' Returns an expression representing an Excel Cell 
    ''' </summary>
    ''' <param name="Row">1-based Row, 1 = Row 1, 0 = Entire Column</param>
    ''' <param name="Col">1-based Column, 1 = Column A, 0 = Entire Row</param>
    ''' <param name="ABSOLUTE">0 = Nothing, 1 = Absolute Column, 2 = Absolute Row, 3 = Absolute Cell</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Excel_Cell(Row As Integer, Col As Integer, Optional ABSOLUTE As Integer = 0) As String

        Dim c As String = Chr(((Col - 1) Mod 26) + 65)
        If Int((Col - 1) / 26) > 0 Then
            c = Chr(Int((Col - 1) / 26) + 64) & c
        End If

        Dim z1 As String = ""
        Dim z2 As String = ""

        If Row > 0 Then
            If ABSOLUTE = 1 Then
                z1 = "$"
            End If
            If ABSOLUTE = 2 Then
                z2 = "$"
            End If
            If ABSOLUTE = 3 Then
                z1 = "$"
                z2 = "$"
            End If
        End If

        Excel_Cell = z1 & c & z2 & IIf(Row > 0, CStr(Row), "")
    End Function
    ''' <summary>
    ''' Returns an expression representing an Excel Cell
    ''' This is a 0-based version of the Excel_Cell function
    ''' </summary>
    ''' <param name="Row">0-based Row, 0 = Row 1, -1 = Entire Column</param>
    ''' <param name="Col">0-based Column, 0 = Column A, -1 = Entire Row</param>
    ''' <param name="ABSOLUTE">0 = Nothing, 1 = Absolute Row, 2 = Absolute Column, 3 = Absolute Cell</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Excel_Cell0(Row As Integer, Col As Integer, Optional ABSOLUTE As Integer = 0) As String
        Return Excel_Cell(Row + 1, Col + 1, ABSOLUTE)
    End Function


    Private Sub btnXLS_LORD_Click(sender As Object, e As EventArgs) Handles btnXLS_LORD.Click


    End Sub

    Private Sub btnXLS_BELK_Click(sender As Object, e As EventArgs) Handles btnXLS_BELK.Click

    End Sub

    Private Sub UltraButton3_Click(sender As Object, e As EventArgs) Handles UltraButton3.Click
        Dim Folder As String = "C:\Users\wjz\source\repos\WinFormsApp1\WinFormsApp1\ADS\"
        Folder = "C:\Users\wjz\VS\AHA\ABS\DataSets\"
        Dim dstITM As New DataSet
        dstITM.ReadXmlSchema(Folder & "WWIMPZITM.XSD")
        dstITM.EnforceConstraints = False
        dstITM.ReadXml(Folder & "WWIMPZITM_SAMPLE.XML")
        grdITM.DataSource = dstITM
        Format_grid(grdITM, dstITM)

        Dim dstPOH As New DataSet
        dstPOH.ReadXmlSchema(Folder & "WWIMPZPOH.XSD")
        dstPOH.EnforceConstraints = False
        dstPOH.ReadXml(Folder & "WWIMPZPOH_SAMPLE.XML")
        grdPOH.DataSource = dstPOH
        Format_grid(grdPOH, dstPOH)

        Dim dstSOH As New DataSet
        dstSOH.ReadXmlSchema(Folder & "WWIMPZSOH.XSD")
        dstSOH.EnforceConstraints = False
        dstSOH.ReadXml(Folder & "WWIMPZSOH_SAMPLE.XML")
        grdSOH.DataSource = dstSOH

        Dim dstRECEIPTS As New DataSet
        dstRECEIPTS.ReadXmlSchema(Folder & "X3RECEIPTS.XSD")
        dstRECEIPTS.EnforceConstraints = False
        dstRECEIPTS.ReadXml(Folder & "X3RECEIPTS_SAMPLE.XML")
        grdRECEIPTS.DataSource = dstRECEIPTS

        Dim dstSHIPMENTS As New DataSet
        'dstSHIPMENTS.ReadXmlSchema(Folder & "X3SHIPMENTS.XSD")
        'dstSHIPMENTS.EnforceConstraints = False
        dstSHIPMENTS.ReadXml(Folder & "X3SHIPMENTS_SAMPLE.XML")
        grdSHIPMENTS.DataSource = dstSHIPMENTS
        ' FORMAT_GRID(grdSHIPMENTS)

        cmdLoadABS.Visible = True
        cmdExport.Visible = True
    End Sub

    Sub Format_grid(grd As UltraWinGrid.UltraGrid, dst As DataSet)
        For b As Integer = 0 To dst.Tables.Count - 1
            Dim tbl As DataTable = dst.Tables(b)
            For Each dc As DataColumn In tbl.Columns
                Dim dt As String = dc.DataType.ToString
                If dt = "System.Integer" Or dt = "System.Decimal" Or dt = "System.Int64" Or dt = "System.Int32" Then
                    With grd.DisplayLayout.Bands(b).Columns(dc.ColumnName)
                        .CellAppearance.TextHAlign = HAlign.Right
                        .Header.Appearance.TextHAlign = HAlign.Right

                        Dim f As String = "#,##0"
                        If dt = "Decimal" Then
                            f = f & ".0000"
                        End If
                        .Format = f
                    End With

                End If
            Next
        Next

    End Sub
    Private Sub cmdLoadABS_Items_Click(sender As Object, e As EventArgs) Handles cmdLoadABS.Click

        Dim tab As String = tabADS.SelectedTab.Text

        Select Case tab

            Case "ITM"

                Dim dst As DataSet = DirectCast(grdITM.DataSource, DataSet)
                dst.Tables("Sales").Rows.Clear()
                Dim t As String = "Item"
                dst.Tables(t).Rows.Clear()

                ASCMAIN1.sql = "Select ICTITEM1.*, TATCNTRY.COUNTRY_CODE2 from ICTITEM1,TATCNTRY"
                ASCMAIN1.sql &= " where TATCNTRY.COUNTRY_CODE(+) = ICTITEM1.COUNTRY_CODE and ICTITEM1.ITEM_CODE IN ('CC003C35USA','CC012C08USA','CC005C02USA','CC016C03USA','CC004C29USA','CC010C11USA','CCZPA132USA','CC001C36USA','CC009C13USA','CC003C37USA','CC006C02USA','CC010C12USA','CC004C31USA','CC017C01USA')"
                For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                    Dim rowITM As DataRow = dst.Tables(t).NewRow
                    With rowITM
                        .Item("StockNumber") = row.Item("ITEM_CODE")
                        .Item("Description1") = row.Item("ITEM_DESC")
                        '.Item("Description2") = row.Item("COLLECTION_CODE")
                        '.Item("Description3") = row.Item("ITEM_CLASS_CODE")
                        Dim ITEM_EAN_CODE As String = row.Item("ITEM_EAN_CODE") & ""
                        Dim ITEM_UPC_CODE As String = row.Item("ITEM_UPC_CODE") & ""
                        Dim EAN_UPC As String = ITEM_EAN_CODE
                        If EAN_UPC = "" Then
                            EAN_UPC = ITEM_UPC_CODE
                        End If
                        .Item("UpcEanCode") = EAN_UPC

                        .Item("ItemWeight") = row.Item("ITEM_WEIGHT")
                        .Item("Category") = "ZFIN" ' row.Item("PROD_CODE")
                        .Item("CountryOfOrigin") = "FR" ' row.Item("COUNTRY_CODE2")
                        .Item("HsCode") = row.Item("HMAT_CODE")
                        .Item("ShelfLifeDays") = Val(row.Item("ITEM_SHELF_LIFE_YRS") & "") * 365
                        .Item("PurchasePrice") = row.Item("ITEM_COST_STD")
                    End With
                    dst.Tables(t).Rows.Add(rowITM)
                    Dim Item_ID As Int64 = Val(rowITM.Item("Item_ID") & "")
                    dst.Tables("Sales").Rows.Add(New Object() {Val(row.Item("ITEM_RETAIL_PRICE") & ""), Item_ID})

                Next

            Case "POH"
                Dim dst As DataSet = DirectCast(grdPOH.DataSource, DataSet)
                dst.Tables("Line").Rows.Clear()
                Dim t As String = "PO"
                dst.Tables(t).Rows.Clear()

                ASCMAIN1.sql = "Select ICTPINV1.*, POTORDR1.WHSE_CODE
                    from POTORDR1,ICTPINV1
                    where POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO and ICTPINV1.PINV_STATUS = 'O' AND ICTPINV1.WHSE_CODE = 'ADS'"
                For Each row As DataRow In ASCDATA1.GetDataTable("").Select("")
                    Dim PINV_NO As String = row.Item("PINV_NO")
                    Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                    Dim INV_DATE As Date = row.Item("INV_DATE")
                    Dim INV_NUM As String = row.Item("INV_NUM")

                    Dim rowPOH As DataRow = dst.Tables(t).NewRow
                    With rowPOH
                        '.Item("Supplier") = row.Item("VEND_CODE")
                        .Item("Supplier") = "MAIN"
                        .Item("ExpectedReceiptDate") = INV_DATE.AddDays(30)
                        .Item("PurchaseOrderNumber") = INV_NUM
                        .Item("InternalReference") = row.Item("PO_ORDER_NO")
                    End With
                    dst.Tables(t).Rows.Add(rowPOH)
                    Dim PO_ID As Int64 = Val(rowPOH.Item("PO_ID") & "")

                    ASCMAIN1.sql = $"Select ICTPINV2.*, POTORDR2.PO_DATE_REQUIRED
                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                        from POTORDR2,ICTITEM1,ICTPINV2
                        where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE
                        and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO
                        and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO
                        AND ICTPINV2.PINV_NO = '{PINV_NO}'"

                    For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                        Dim ITEM_CODE As String = row2.Item("ITEM_CODE")
                        Dim PINV_QTY As Int32 = Val(row2.Item("PINV_QTY") & "")
                        Dim PINV_LNO As Int32 = Val(row2.Item("PINV_LNO") & "")
                        Dim PO_ORDER_LNO As Int32 = Val(row2.Item("PO_ORDER_LNO") & "")

                        'ASCMAIN1.sql = $"SELECT ICTPINV1.PINV_REF_INV
                        '    FROM ICTPINV1,ICTPINV2 WHERE ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                        '    AND ICTPINV1.PINV_STATUS= 'O' AND ICTPINV1.REVERSED_BY_PINV_NO IS NULL
                        '    AND ICTPINV2.PO_ORDER_NO = '{PO_ORDER_NO}' AND ICTPINV2.PO_ORDER_LNO = {CStr(PO_ORDER_LNO)}"
                        'Dim rowPINV As DataRow = ASCDATA1.GetDataRow

                        Dim CustomerOrderRef As String = CStr(PO_ORDER_LNO)
                        'If rowPINV IsNot Nothing Then
                        '    CustomerOrderRef = rowPINV.Item("PINV_REF_INV")
                        'End If


                        If row2.Item("PO_DATE_REQUIRED") & "" <> "" Then
                            'CustomerOrderRef = "Date Req " & Format(row2.Item("PO_DATE_REQUIRED"), "MM/dd/yyyy")
                            If rowPOH.Item("ExpectedReceiptDate") & "" = "" Then
                                rowPOH.Item("ExpectedReceiptDate") = row2.Item("PO_DATE_REQUIRED")
                            End If
                        End If

                        Dim rowLine As DataRow = dst.Tables("Line").NewRow
                        With rowLine
                            .Item("StockNumber") = ITEM_CODE
                            .Item("Quantity") = PINV_QTY
                            .Item("CustomerOrderRef") = CustomerOrderRef
                            .Item("EDISize") = row2.Item("NRF_SIZE_CODE")
                            .Item("EDIColor") = row2.Item("NRF_COLOR_CODE")
                            .Item("PO_ID") = PO_ID
                        End With
                        dst.Tables("Line").Rows.Add(rowLine)
                    Next
                Next



            Case "SOH"



                Dim dst As DataSet = DirectCast(grdSOH.DataSource, DataSet)
                dst.Tables("Line").Rows.Clear()
                Dim t As String = "Order"
                dst.Tables(t).Rows.Clear()


                ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_ARRIVAL_DATE
                    , SOTORDR1.CUST_CODE, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO, SOTORDR1.ORDR_DEPT, SOTORDR1.CUST_BILL_TO_CUST
                    , SOTSHIP1.SHIP_BOL_NO
                    from SOTPICK1,SOTORDR1,SOTSHIP1
                    where SOTPICK1.PICK_STATUS = 'P'
                      and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                      and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                      and ROWNUM < 1000"
                '      and SOTSHIP1.WHSE_CODE = 'CLA'"
                '      and SOTSHIP1.WHSE_CODE = 'ADS'"
                ' IN FUTIRE ADD PICK_PRINTED IS NULL AND THEN POPULATE PICK_PRINTED

                For Each row As DataRow In ASCDATA1.GetDataTable().Select("")

                    Dim PICK_NO As String = row.Item("PICK_NO")
                    Dim ORDR_NO As String = row.Item("ORDR_NO")
                    Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                    Dim CUST_DC_NO As String = row.Item("CUST_DC_NO") & ""

                    Dim CUST_BILL_TO_CUST As String = row.Item("CUST_BILL_TO_CUST") & ""
                    If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE

                    ASCMAIN1.sql = $"Select * from ARTCUST1 where CUST_CODE = '{CUST_BILL_TO_CUST}'"
                    Dim rowSOTORDR5_BT As DataRow = ASCDATA1.GetDataRow
                    ASCMAIN1.sql = $"Select * from SOTORDR5 where ORDR_NO = '{ORDR_NO}' and CUST_ADDR_TYPE = 'ST'"
                    Dim rowSOTORDR5_ST As DataRow = ASCDATA1.GetDataRow

                    Dim rowSOH As DataRow = dst.Tables(t).NewRow
                    With rowSOH
                        '.Item("SalesSite") = ORDR_NO
                        '.Item("AutoCreateIfMissing") = ORDR_NO
                        '.Item("EDISenderID") = ORDR_NO
                        '.Item("EDIReceiverID") = ORDR_NO
                        .Item("SalesOrderNumber") = PICK_NO
                        .Item("ExternalOrderNumber") = SHIP_BOL_NO
                        .Item("SoldToCustomer") = row.Item("CUST_CODE")
                        .Item("SalesOrderDate") = row.Item("ORDR_DATE")
                        .Item("CustomerPONumber") = row.Item("ORDR_CUST_PO")
                        .Item("Department") = row.Item("ORDR_DEPT")
                        ' .Item("Division") = row.Item("ORDR_DEPT")
                        .Item("Currency") = "USD"
                        .Item("Carrier") = "TBD"
                        ' .Item("ThirdPartyFrtAcct") = row.Item("ORDR_DEPT")
                        ' .Item("FreightInvoicing") = row.Item("ORDR_DEPT")
                        '.Item("CustomerFreight") = row.Item("ORDR_DEPT")
                        '.Item("PackingList") = row.Item("ORDR_DEPT")
                        .Item("RequestedDeliveryDat") = row.Item("ORDR_ARRIVAL_DATE")
                        .Item("StartShipDate") = row.Item("ORDR_SHIP_DATE")
                        .Item("ShipDate") = row.Item("ORDR_SHIP_DATE")


                        .Item("BillToCode") = CUST_BILL_TO_CUST
                        .Item("BillToName") = rowSOTORDR5_BT.Item("CUST_NAME")
                        .Item("BillToAddress1") = rowSOTORDR5_BT.Item("CUST_ADDR1")
                        .Item("BillToAddress2") = rowSOTORDR5_BT.Item("CUST_ADDR2")
                        .Item("BillToAddress3") = rowSOTORDR5_BT.Item("CUST_ADDR3")
                        .Item("BillToCity") = rowSOTORDR5_BT.Item("CUST_CITY")
                        .Item("BillToState") = rowSOTORDR5_BT.Item("CUST_STATE")
                        .Item("BillToZipcode") = rowSOTORDR5_BT.Item("CUST_ZIP_CODE")
                        .Item("BillToCountry") = rowSOTORDR5_BT.Item("CUST_COUNTRY")

                        .Item("ShipToCode") = CUST_STORE_NO
                        .Item("ShipToName") = rowSOTORDR5_ST.Item("CUST_NAME")
                        .Item("ShipToAddress1") = rowSOTORDR5_ST.Item("CUST_ADDR1")
                        .Item("ShipToAddress2") = rowSOTORDR5_ST.Item("CUST_ADDR2")
                        .Item("ShipToAddress3") = rowSOTORDR5_ST.Item("CUST_ADDR3")
                        .Item("ShipToCity") = rowSOTORDR5_ST.Item("CUST_CITY")
                        .Item("ShipToState") = rowSOTORDR5_ST.Item("CUST_STATE")
                        .Item("ShipToZipcode") = rowSOTORDR5_ST.Item("CUST_ZIP_CODE")
                        .Item("ShipToCountry") = rowSOTORDR5_ST.Item("CUST_COUNTRY")

                        .Item("DcCode") = CUST_DC_NO
                        '.Item("TaxCode") = CUST_STORE_NO
                        '.Item("UctTaxAmount") = CUST_STORE_NO
                        '.Item("UctTaxRate") = CUST_STORE_NO
                        '.Item("DiscountPercent") = CUST_STORE_NO
                        '.Item("DiscountAmount") = CUST_STORE_NO

                    End With
                    dst.Tables(t).Rows.Add(rowSOH)
                    Dim Order_ID As Int64 = Val(rowSOH.Item("Order_ID") & "")


                    ASCMAIN1.sql = $"Select SOTPICK2.*
                        , SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.ITEM_RETAIL_PRICE
                        , ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                        from SOTPICK2,SOTORDR2,ICTITEM1
                        where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE
                            AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO
                            AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO
                        And SOTPICK2.PICK_NO = '{PICK_NO}'"
                    For Each row2 As DataRow In ASCDATA1.GetDataTable().Select("")
                        Dim ITEM_CODE As String = row2.Item("ITEM_CODE")
                        Dim rowLine As DataRow = dst.Tables("Line").NewRow
                        With rowLine
                            ' .Item("ShipToCode") = Val(row2.Item("ORDR_LNO") & "")
                            .Item("LineNumber") = Val(row2.Item("PICK_LNO") & "")
                            .Item("ItemReferenceNumber") = row2.Item("ITEM_CODE")
                            '.Item("CustomerItemRef") = row2.Item("ITEM_CODE")
                            Dim ITEM_EAN_CODE As String = row2.Item("ITEM_EAN_CODE") & ""
                            Dim ITEM_UPC_CODE As String = row2.Item("ITEM_UPC_CODE") & ""
                            Dim EAN_UPC As String = ITEM_EAN_CODE
                            If EAN_UPC = "" Then
                                EAN_UPC = ITEM_UPC_CODE
                            End If
                            .Item("UpcEanCode") = EAN_UPC
                            .Item("EDISize") = row2.Item("NRF_SIZE_CODE")
                            .Item("EDIColor") = row2.Item("NRF_COLOR_CODE")
                            .Item("ItemDescription") = row2.Item("ITEM_DESC")
                            .Item("ItemQuantity") = Val(row2.Item("PICK_QTY") & "")
                            .Item("UnitOfMeasure") = "EA"
                            '.Item("UomConversionQty") = "EA"
                            .Item("ItemGrossPrice") = Val(row2.Item("ORDR_UNIT_PRICE") & "")
                            .Item("ItemCustGrossPrice") = Val(row2.Item("ITEM_RETAIL_PRICE") & "")
                            .Item("ItemDiscount1") = (Val(row2.Item("ITEM_RETAIL_PRICE") & "") - Val(row2.Item("ORDR_UNIT_PRICE") & "")) * Val(row2.Item("PICK_QTY") & "")
                            .Item("Order_ID") = Order_ID
                        End With
                        dst.Tables("Line").Rows.Add(rowLine)
                    Next
                Next

        End Select

    End Sub

    Private Sub cmdExport_Click(sender As Object, e As EventArgs) Handles cmdExport.Click

        'ALTER TABLE TATSSHK1 MODIFY SSH_APP_USERNAME VARCHAR2(20);
        'ALTER TABLE TATSSHK1 MODIFY SSH_APP_PASSWORD VARCHAR2(20);
        'ALTER TABLE TATSSHK1 ADD SSH_APP_PORT NUMBER (6,0);
        'INSERT INTO TATSSHK1 (SSH_APP_CODE,SSH_APP_DESC,SSH_APP_USERNAME,SSH_APP_PASSWORD,SSH_APP_PARTNER_URI_PROD,SSH_APP_FOLDER_GET,SSH_APP_FOLDER_PUT, SSH_APP_PORT) VALUES ('ADS','ADS','cust_interparfu','cust_interparfu','ads-live.dyndns.org', 'out','in',6722)

        Dim dts As String = Format(Now, "yyyyMMdd_HHmmss")
        dts = ASCMAIN1.Next_Control_No("SOTSHIP1.LP_XNO") ' check with ed
        Dim XSD_FILENAME As String = ""
        Dim tab As String = tabADS.SelectedTab.Text
        Dim dst As DataSet = Nothing
        Select Case tab

            Case "ITM"
                XSD_FILENAME = $"WWIMPZITM_{dts}.XML"
                dst = DirectCast(grdITM.DataSource, DataSet)

            Case "POH"
                XSD_FILENAME = $"WWIMPZPOH_{dts}.XML"
                dst = DirectCast(grdPOH.DataSource, DataSet)
        End Select

        If XSD_FILENAME <> "" Then
            ' dst.WriteXml(XSD_FILENAME, XmlWriteMode.WriteSchema)
            dst.WriteXml(XSD_FILENAME)

            If chksftp.Checked Then
                sftp_put(Nothing, "ADS", True, XSD_FILENAME, XSD_FILENAME)

            End If
        End If
    End Sub

    Function sftp_put(
        frmASFBASE0 As ASFBASE0,
        SSH_APP_CODE As String,
        production As Boolean,
        FILENAME_LOCAL As String,
        FILENAME_REMOTE As String) As Boolean

        Dim rowTATSSHK1 As DataRow = Nothing ' frmASFBASE0.LookUp("TATSSHK1", SSH_APP_CODE)

        ' SHOULD BE USING EXP COMPANY FOR A&E
        rowTATSSHK1 = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")

        Dim SSH_APP_USERNAME As String = rowTATSSHK1.Item("SSH_APP_USERNAME") & ""
        Dim SSH_APP_PASSWORD As String = rowTATSSHK1.Item("SSH_APP_PASSWORD") & ""
        Dim SSH_APP_FOLDER_PUT As String = rowTATSSHK1.Item("SSH_APP_FOLDER_PUT") & ""
        'If SSH_APP_FOLDER_PUT.EndsWith("\") Then

        'End If

        Dim SSH_APP_PARTNER_URI As String = ""
        Dim SSH_APP_PARTNER_PUBKEY As String = ""
        If production Then
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & ""
        Else
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & ""
        End If

        Dim SSH_APP_PORT As Integer = Val(rowTATSSHK1.Item("SSH_APP_PORT") & "")
        If SSH_APP_PORT = 0 Then
            SSH_APP_PORT = 22
        End If

        Dim success = False
        Dim sftp As New nsoftware.IPWorksSSH.Sftp
        theLog = ""

        AddHandler sftp.OnSSHServerAuthentication, AddressOf SSHServerAuthentication
        AddHandler sftp.OnSSHStatus, AddressOf SSHStatus

        sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")


        If SSH_APP_CODE = "EXP" Then

        Else

            sftp.SSHUser = SSH_APP_USERNAME

            If SSH_APP_PASSWORD <> "" Then
                sftp.SSHAuthMode = SftpSSHAuthModes.amPassword
                sftp.SSHPassword = SSH_APP_PASSWORD
            Else
                sftp.SSHAuthMode = SftpSSHAuthModes.amPublicKey
                'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
                'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")

                If ASCMAIN1.Running_in_VS Then
                    Stop
                    sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\VS\AHA\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
                    'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
                Else
                    ' sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
                    'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
                    Dim ssh_file As String = ASCMAIN1.Folders("SharedRoot") & "Archive\INT\JPMC\JPMC_SSH_pvt.ppk"
                    sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, ssh_file, "0ff1c3INT", "*")
                End If

            End If

            Try

                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If



                If ASCMAIN1.CLIENT = "INT" Then
                    If SSH_APP_CODE = "JPMC" Then
                        sftp.SSHEncryptionAlgorithms = "aes128-ctr,aes192-ctr,aes256-ctr"
                        sftp.Config("LogSSHPackets=True")
                        If ASCMAIN1.USER_ID = "wjz" Then MsgBox(sftp.Config("LogSSHPackets"))
                    Else
                        ' COWORX DOES NOT SUPPORT NEW ENCRYPTION
                    End If
                    If SSH_APP_CODE = "ADS" Then
                        SSH_APP_PORT = 6722
                    End If
                End If

                sftp.SSHHost = SSH_APP_PARTNER_URI
                sftp.SSHLogon(SSH_APP_PARTNER_URI, SSH_APP_PORT)
                success = True

                sftp.LocalFile = FILENAME_LOCAL
                sftp.RemotePath = SSH_APP_FOLDER_PUT

                sftp.RemoteFile = FILENAME_REMOTE
                sftp.Upload()

            Catch ex As Exception
                theLog &= ex.Message
                Dim filename As String = Format(Now, "yyyyMMddhhhhss")
                System.IO.File.WriteAllText(ASCMAIN1.Folders("Work") & filename & ".log", theLog)
                MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If sftp.Connected = True Then
                    sftp.SSHLogoff()
                End If
            End Try

        End If

        If sftp.Connected = True Then
            sftp.SSHLogoff()
        End If

        Return success

    End Function

    Public Shared Function sftp_get(
        frmASFBASE0 As ASFBASE0,
        SSH_APP_CODE As String,
        production As Boolean,
        FILENAME_LOCAL As String,
        FILENAME_REMOTE As String) As List(Of String)

        Dim FILENAMEs As New List(Of String)

        Dim rowTATSSHK1 As DataRow = frmASFBASE0.LookUp("TATSSHK1", SSH_APP_CODE)

        Dim SSH_APP_USERNAME As String = rowTATSSHK1.Item("SSH_APP_USERNAME") & ""
        Dim SSH_APP_PASSWORD As String = rowTATSSHK1.Item("SSH_APP_PASSWORD") & ""
        Dim SSH_APP_FOLDER_GET As String = rowTATSSHK1.Item("SSH_APP_FOLDER_GET") & ""
        'If SSH_APP_FOLDER_PUT.EndsWith("\") Then

        'End If

        Dim SSH_APP_PARTNER_URI As String = ""
        Dim SSH_APP_PARTNER_PUBKEY As String = ""
        If production Then
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_PROD") & ""
        Else
            SSH_APP_PARTNER_URI = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_TEST") & ""
            SSH_APP_PARTNER_PUBKEY = rowTATSSHK1.Item("SSH_APP_PARTNER_PUBKEY_TEST") & ""
        End If

        Dim success = False
        Dim sftp As New nsoftware.IPWorksSSH.Sftp
        AddHandler sftp.OnSSHServerAuthentication, AddressOf SSHServerAuthentication

        sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")

        sftp.SSHUser = SSH_APP_USERNAME

        If SSH_APP_PASSWORD <> "" Then
            sftp.SSHAuthMode = SftpSSHAuthModes.amPassword
            sftp.SSHPassword = SSH_APP_PASSWORD
        Else
            sftp.SSHAuthMode = SftpSSHAuthModes.amPublicKey
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_SSH_pvt.ppk", "0ff1c3INT", "*")
            'sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, "S:\INT\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
            sftp.SSHCert = New Certificate(CertStoreTypes.cstPPKFile, ASCMAIN1.Folders("SharedRoot") & "\Archive\INT\JPMC\JPMC_IPLB_pvt.asc", "0ff1c3INT", "*")
        End If

        Try

            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If

            sftp.SSHHost = SSH_APP_PARTNER_URI
            sftp.SSHLogon(SSH_APP_PARTNER_URI, 22)
            success = True
            sftp.RemotePath = "/" & SSH_APP_FOLDER_GET

            sftp.ListDirectory()
            For Each s As nsoftware.IPWorksSSH.DirEntry In sftp.DirList
                sftp.RemoteFile = s.FileName
                If Not s.IsDir Then
                    ASCMAIN1.Progress("-", s.FileName)
                    sftp.LocalFile = FILENAME_LOCAL & s.FileName
                    sftp.Download()
                    '  sftp.RenameFile(FILENAME_LOCAL & "\Archive\" & s.FileName)

                    sftp.DeleteFile(s.FileName)
                    FILENAMEs.Add(FILENAME_LOCAL & s.FileName)
                End If
            Next

            sftp.SSHLogoff()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Secure ftp Setup", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If sftp.Connected = True Then
                sftp.SSHLogoff()
            End If
        End Try

        If sftp.Connected = True Then
            sftp.SSHLogoff()
        End If

        Return FILENAMEs ' success
    End Function


    Public Shared Sub SSHServerAuthentication(sender As Object, e As SftpSSHServerAuthenticationEventArgs)

        e.Accept = True
    End Sub

    Sub SSHStatus(sender As Object, e As SftpSSHStatusEventArgs)

        ' MsgBox(e.Message, MsgBoxStyle.OkOnly, "SSHStatus Messages")
        theLog &= e.Message & vbCrLf

    End Sub

    Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function

    Private Sub btnJPMC_Click(sender As Object, e As EventArgs) Handles btnJPMC.Click
        Dim Folder As String = "C:\vs\aha\work\"
        Dim dstChase As New DataSet
        dstChase.ReadXmlSchema(Folder & "chase.XSD")
        dstChase.EnforceConstraints = False
        'dstChase.ReadXml(Folder & "WWIMPZITM_SAMPLE.XML")
        grdChase.DataSource = dstChase
        'Format_grid(grdITM, dstChase)


        If False Then
            Dim dst As DataSet = DirectCast(grdPOH.DataSource, DataSet)
            dst.Tables("Line").Rows.Clear()
            Dim t As String = "PO"
            dst.Tables(t).Rows.Clear()

            ASCMAIN1.sql = "Select ICTPINV1.*, POTORDR1.WHSE_CODE
                    from POTORDR1,ICTPINV1
                    where POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO and ICTPINV1.PINV_STATUS = 'O' AND ICTPINV1.WHSE_CODE = 'ADS'"
            For Each row As DataRow In ASCDATA1.GetDataTable("").Select("")
                Dim PINV_NO As String = row.Item("PINV_NO")
                Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                Dim INV_DATE As Date = row.Item("INV_DATE")
                Dim INV_NUM As String = row.Item("INV_NUM")

                Dim rowPOH As DataRow = dst.Tables(t).NewRow
                With rowPOH
                    '.Item("Supplier") = row.Item("VEND_CODE")
                    .Item("Supplier") = "MAIN"
                    .Item("ExpectedReceiptDate") = INV_DATE.AddDays(30)
                    .Item("PurchaseOrderNumber") = INV_NUM
                    .Item("InternalReference") = row.Item("PO_ORDER_NO")
                End With
                dst.Tables(t).Rows.Add(rowPOH)
                Dim PO_ID As Int64 = Val(rowPOH.Item("PO_ID") & "")

                ASCMAIN1.sql = $"Select ICTPINV2.*, POTORDR2.PO_DATE_REQUIRED
                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                        from POTORDR2,ICTITEM1,ICTPINV2
                        where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE
                        and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO
                        and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO
                        AND ICTPINV2.PINV_NO = '{PINV_NO}'"

                For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                    Dim ITEM_CODE As String = row2.Item("ITEM_CODE")
                    Dim PINV_QTY As Int32 = Val(row2.Item("PINV_QTY") & "")
                    Dim PINV_LNO As Int32 = Val(row2.Item("PINV_LNO") & "")
                    Dim PO_ORDER_LNO As Int32 = Val(row2.Item("PO_ORDER_LNO") & "")

                    'ASCMAIN1.sql = $"SELECT ICTPINV1.PINV_REF_INV
                    '    FROM ICTPINV1,ICTPINV2 WHERE ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                    '    AND ICTPINV1.PINV_STATUS= 'O' AND ICTPINV1.REVERSED_BY_PINV_NO IS NULL
                    '    AND ICTPINV2.PO_ORDER_NO = '{PO_ORDER_NO}' AND ICTPINV2.PO_ORDER_LNO = {CStr(PO_ORDER_LNO)}"
                    'Dim rowPINV As DataRow = ASCDATA1.GetDataRow

                    Dim CustomerOrderRef As String = CStr(PO_ORDER_LNO)
                    'If rowPINV IsNot Nothing Then
                    '    CustomerOrderRef = rowPINV.Item("PINV_REF_INV")
                    'End If


                    If row2.Item("PO_DATE_REQUIRED") & "" <> "" Then
                        'CustomerOrderRef = "Date Req " & Format(row2.Item("PO_DATE_REQUIRED"), "MM/dd/yyyy")
                        If rowPOH.Item("ExpectedReceiptDate") & "" = "" Then
                            rowPOH.Item("ExpectedReceiptDate") = row2.Item("PO_DATE_REQUIRED")
                        End If
                    End If

                    Dim rowLine As DataRow = dst.Tables("Line").NewRow
                    With rowLine
                        .Item("StockNumber") = ITEM_CODE
                        .Item("Quantity") = PINV_QTY
                        .Item("CustomerOrderRef") = CustomerOrderRef
                        .Item("EDISize") = row2.Item("NRF_SIZE_CODE")
                        .Item("EDIColor") = row2.Item("NRF_COLOR_CODE")
                        .Item("PO_ID") = PO_ID
                    End With
                    dst.Tables("Line").Rows.Add(rowLine)
                Next
            Next

        End If
    End Sub


    Private Sub Load_ST()

        Dim TBL As New DataTable
        With TBL.Columns
            .Add("KEY")
            Dim SQL As String = ""
            For I As Integer = 0 To 12
                Dim C As String = "D" & Format(I, "00")
                If I <> 0 Then SQL &= $"+{C}"
                .Add(C, GetType(System.Int32))
            Next
            TBL.Columns("D00").Expression = Mid(SQL, 2)
            SQL = ""
            For I As Integer = 0 To 12
                Dim C As String = "E" & Format(I, "00")
                If I <> 0 Then SQL &= $"+{C}"
                .Add(C, GetType(System.Int32))
            Next
            TBL.Columns("E00").Expression = Mid(SQL, 2)
        End With

        For R As Integer = 1 To 10
            Dim RR As String = Format(R, "00")
            Dim row As DataRow = TBL.NewRow
            row.Item("KEY") = RR
            TBL.Rows.Add(row)

            Static Generator As System.Random = New System.Random()

            For D As Integer = 1 To 10
                Dim V As Int32 = Generator.Next(0, 100)
                Dim C As String = "D" & Format(D, "00")
                row.Item(C) = V
            Next

            For E As Integer = 1 To 10
                Dim V As Int32 = Generator.Next(0, 100)
                Dim C As String = "E" & Format(E, "00")
                row.Item(C) = V
            Next
        Next
        grdST.DataSource = TBL

        grdST.DisplayLayout.Bands(0).LevelCount = 2
        With grdST.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup
            G = .Groups.Add("KEY")
            .Columns("KEY").Group = G
            For Z As Integer = 0 To 12
                G = .Groups.Add("G" & Format(Z, "00"))
                .Columns("D" & Format(Z, "00")).Group = G
                .Columns("D" & Format(Z, "00")).Level = 0
                .Columns("E" & Format(Z, "00")).Group = G
                .Columns("E" & Format(Z, "00")).Level = 1
            Next
        End With

        grdST.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdST.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdST.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdST.DisplayLayout.Bands(0).ColHeadersVisible = False
        grdST.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdST.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.Empty
        grdST.DisplayLayout.Override.ActiveRowAppearance.ForeColor = System.Drawing.Color.Empty
        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(5)
    End Sub

    Private Sub UltraButton4_Click(sender As Object, e As EventArgs) Handles UltraButton4.Click
        Load_ST()
    End Sub

    Private Sub grdST_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdST.InitializeLayout

    End Sub

    Private Sub grdST_MouseUp(sender As Object, e As MouseEventArgs) Handles grdST.MouseUp
        If grdST.Selected.Cells.Count = 0 Then
            lblST.Text = "Total Selected Cells"
        Else

            Dim L As Integer = grdST.DisplayLayout.Bands(0).LevelCount

            Dim T() As Int32
            ReDim T(L - 1)
            For Each cell As UltraWinGrid.UltraGridCell In grdST.Selected.Cells
                Dim LX As Integer = cell.Column.Level
                T(LX) += Val(cell.Value & "")
            Next
            Dim LZ As String = ""
            For II As Integer = 0 To L - 1
                LZ &= ", " & Format(T(II), "#,##0")
            Next
            lblST.Text = $"Total Selected Cells ({grdST.Selected.Cells.Count}) = {Mid(LZ, 3)}"
        End If
    End Sub


    Private Async Sub btnJWT_Click(sender As Object, e As EventArgs) Handles btnJWT.Click
        ' Stop
        ' https://stackoverflow.com/questions/32716174/call-and-consume-web-api-in-winform-using-c-net

        Dim client As HttpClient = New HttpClient()
        Dim url As String = "https://absapi.absolution1.com/api/Home/"
        ' url = "http://localhost:1977/"
        url = "https://absapi.absolution1.com/"
        'url = "http://localhost:1642/"

        ' Put the following code where you want to initialize the class
        ' It can be the static constructor Or a one-time initializer
        ' client.BaseAddress = New Uri("http://localhost:4354/api/")
        client.BaseAddress = New Uri(url)
        client.DefaultRequestHeaders.Accept.Clear()
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

        ' Assuming http//localhost:4354/api/ as BaseAddress 
        ' Dim response As Object = Await client.GetAsync("persons")


        ' Assuming http//localhost:4354/api/ as BaseAddress 
        'var product = New Product() {Name = "P1", Price = 100, Category = "C1"};
        'var response = await client.PostAsJsonAsync("products", product);

        '        [
        '{
        '"Name": "Name 1",
        '"Email": "Email 1",
        '"Address": "Address 1"
        '},
        '{
        '"Name": "Name 2",
        '"Email": "Email 2",
        '"Address": "Address 2"
        '},
        '{
        '"Name": "Name 3",
        '"Email": "Email 3",
        '"Address": "Address 3"
        '}
        ']

        ' Assuming http//localhost4354/api/ as BaseAddress 
        'Dim response As Object = Await client.GetStringAsync("persons")
        'Dim Data = JsonConvert.DeserializeObject(Of List(Of person))(response)
        '' this.productBindingSource.DataSource = Data; 
        Dim env As String = "PROD"
        If chkUseUAT.Checked Then env = "UAT"

        Dim response As String = Await client.GetStringAsync($"GetJPMC_JWT/{env}")

        ' .GetAsync(String.Format("api/products/id={0}&type={1}", param.Id.Value, param.Id.Type)).Result;

        txtJWT.Text = Mid(response, 2, response.Length - 2)
        'If you have used methods Like GetAsync Or PostAsJsonAsync And you have an HttpResponseMessage Then you can use ReadAsAsync, ReadAsByteArrayAsync, ReadAsStreamAsync, `ReadAsStringAsync, For example:

        ' Assuming http//localhost4354/api/ as BaseAddress 
        ' var response = Await client.GetAsync("products");
        'var Data = Await response.Content.ReadAsAsync < IEnumerable < Product >> ();
        'this.productBindingSource.DataSource = Data;

        ' Stop
    End Sub

    Private Async Sub btnOAuth_Click(sender As Object, e As EventArgs) Handles btnOAuth.Click

        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC As String = "https://idauatg2.jpmorganchase.com/adfs/oauth2/" ' UAT
        urlJPMC = "https://idag2.jpmorganchase.com/adfs/oauth2/token/" ' PROD
        urlJPMC = "https://idag2.jpmorganchase.com/adfs/oauth2/" ' PROD
        clientJPMC.BaseAddress = New Uri(urlJPMC)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))

        'Dim req As OAuthRequest = New OAuthRequest() With {
        '    .grant_type = "client_credentials",
        '    .client_id = lblCLIENT_ID.Text,
        '    .resource = "https://apigeeproductProd.jpmchase.net",
        '    .client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        '    .client_assertion = txtJWT.Text}

        Dim CLIENT_ID As String = lblCLIENT_ID_UAT.Text ' UAT
        CLIENT_ID = lblCLIENT_ID.Text ' PROD

        Dim JWT As String = txtJWT.Text ' UAT & PROD

        ' This is the POST body
        Dim reqkv As New List(Of KeyValuePair(Of String, String))
        reqkv.Add(New KeyValuePair(Of String, String)("grant_type", "client_credentials"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_id", CLIENT_ID))
        reqkv.Add(New KeyValuePair(Of String, String)("resource", "https://apigeeproductProd.jpmchase.net"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_assertion", JWT))

        Dim q As HttpContent = New FormUrlEncodedContent(reqkv)

        Dim res As HttpResponseMessage = Await clientJPMC.PostAsync("token", q)
        Dim body As String = Await res.Content.ReadAsStringAsync()
        Dim c As OAuthResponse = JsonConvert.DeserializeObject(Of OAuthResponse)(body)

        txtOAuth.Text = c.access_token

        ' how to post to upload a file
        ' https://stackoverflow.com/questions/67050528/c-sharp-httpclient-post-request-with-custom-headers-sends-incorrect-content-type

    End Sub

    Public Class OAuthResponse
        Public access_token As String
        Public token_type As String
        Public expires_in As Int32
    End Class

    Private Async Sub btnBalances_Click(sender As Object, e As EventArgs) Handles btnBalances.Click

        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC As String = "https://openbankinguat.jpmorgan.com/accessapi/"
        If Not chkUseUAT.Checked Then
            urlJPMC = "https://openbanking.jpmorgan.com/accessapi/"
        End If
        clientJPMC.BaseAddress = New Uri(urlJPMC)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))
        clientJPMC.DefaultRequestHeaders.Add("Authorization", "Bearer " & txtOAuth.Text)

        'Dim reqkv3 As New List(Of KeyValuePair(Of String, Object))
        'reqkv3.Add(New KeyValuePair(Of String, Object)("accountId", "000000004045701"))

        ' This is the POST body
        'Dim reqkv As New List(Of KeyValuePair(Of String, Object))
        'reqkv.Add(New KeyValuePair(Of String, Object)("relativeDateType", "CURRENT_DAY"))
        'reqkv.Add(New KeyValuePair(Of String, Object)("accountList", reqkv3))

        Dim accountId As String = "000000004045701"
        If Not chkUseUAT.Checked Then
            accountId = "000000899558928" ' WIRE
            'accountId = "000000899558985" ' CHASE
            'accountId = "000000899558928,000000899558985" ' BOTH
        End If

        Dim req2 As New JPMC_Balances_Request2 With {.accountId = accountId}
        Dim req3 As New List(Of JPMC_Balances_Request2)
        req3.Add(req2)
        'Dim relativeDateType As String = ""
        'Dim req1 As New JPMC_Balances_Request With {.relativeDateType = "CURRENT_DAY", .accountList = req3}
        'Dim req1 As New JPMC_Balances_Request With {.relativeDateType = "PRIOR_DAY", .accountList = req3}

        Dim startDate As String = Format(Now.AddDays(-31), "yyyy-MM-dd")
        Dim endDate As String = Format(Now.AddDays(-1), "yyyy-MM-dd")
        'd.Add("startDate", $"{startDate}")
        'd.Add("endDate", $"{endDate}")

        Dim req1 As New JPMC_Balances_Request With {.startDate = startDate, .endDate = endDate, .accountList = req3}

        'Dim q As HttpContent = New FormUrlEncodedContent(req1)
        'Dim res As HttpResponseMessage = Await clientJPMC.PostAsync("balance", q)
        'Dim body As String = Await res.Content.ReadAsStringAsync()
        'Dim c As OAuthResponse = JsonConvert.DeserializeObject(Of OAuthResponse)(body)


        Dim res = Await clientJPMC.PostAsJsonAsync("balance", req1)

        Dim body As String = Await res.Content.ReadAsStringAsync()
        Dim c As JPMC_Balances_Response = JsonConvert.DeserializeObject(Of JPMC_Balances_Response)(body)


        txtBalances.Text = body

        Dim cc As New List(Of JPMC_Balances_Response)
        cc.Add(c)
        grdBalances.DataSource = cc
    End Sub

    Public Class JPMC_Balances_Request
        Public Property relativeDateType As String
        Public Property startDate As String
        Public Property endDate As String
        Public Property accountList As List(Of JPMC_Balances_Request2)
    End Class

    Public Class JPMC_Balances_Request2
        Public Property accountId As String
    End Class

    Public Class AccountList
        Public Property accountId As String
        Public Property accountName As String
        Public Property branchId As String
        Public Property bankId As String
        Public Property bankName As String
        Public Property currency As Currency
        Public Property balanceList As List(Of BalanceList)
    End Class

    Public Class BalanceList
        Public Property asOfDate As String
        Public Property recordTimestamp As DateTime
        Public Property currentDay As Boolean
        Public Property openingAvailableAmount As Double
        Public Property openingLedgerAmount As Double
        Public Property endingAvailableAmount As Double
        Public Property endingLedgerAmount As Double
    End Class

    Public Class Currency
        Public Property code As String
        Public Property currencySequence As Integer
        Public Property decimalLocation As Integer
        Public Property description As String
    End Class

    Public Class JPMC_Balances_Response
        Public Property accountList As List(Of AccountList)
    End Class


    Private Async Sub btnTransactions_Click(sender As Object, e As EventArgs) Handles btnTransactions.Click

        'Stop
        'https://openbankinguat.jpmorgan.com/tsapi/v3/transactions?pageNumber=1&accountIds=000000010013324&relativeDateType=PRIOR_DAY

        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC As String = "https://openbankinguat.jpmorgan.com/tsapi/v3/"
        If Not chkUseUAT.Checked Then
            urlJPMC = "https://openbanking.jpmorgan.com/tsapi/v3/"
        End If
        clientJPMC.BaseAddress = New Uri(urlJPMC)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))
        clientJPMC.DefaultRequestHeaders.Add("Authorization", "Bearer " & txtOAuth.Text)


        ' https://stackoverflow.com/questions/56823997/add-parameters-to-httpclient

        'Dim req = New HttpRequestMessage(HttpMethod.Get, "https://login.microsoftonline.com/0475dfa7-xxxxxxxx-896cf5e31efc/oauth2/token")
        'req.Headers.Add("Referer", "login.microsoftonline.com")
        'req.Headers.Add("Accept", "application/x-www-form-urlencoded")
        ''req.Headers.Add("Content-Type", "application/x-www-form-urlencoded")

        ' This Is the important part

        Dim d As New Dictionary(Of String, String)
        d.Add("pageNumber", "1")

        Dim accountId As String = "000000010013324"
        If Not chkUseUAT.Checked Then
            accountId = "000000899558928" ' WIRE
            accountId = "000000899558985" ' CHASE
            accountId = "000000899558928,000000899558985" ' BOTH
        End If

        d.Add("accountIds", accountId)

        'd.Add("relativeDateType", "PRIOR_DAY")
        'grdTransactions.Text = $"Transactions for PRIOR_DAY"
        'd.Add("relativeDateType", "CURRENT_DAY")
        'grdTransactions.Text = $"Transactions for CURRENT_DAY"

        Dim startDate As String = Format(Now.AddDays(-25), "yyyy-MM-dd")
        Dim endDate As String = Format(Now.AddDays(-21), "yyyy-MM-dd")
        'd.Add("startDate", $"startDate={startDate}")
        'd.Add("endDate", $"endDate={endDate}")
        d.Add("startDate", $"{startDate}")
        d.Add("endDate", $"{endDate}")
        grdTransactions.Text = $"Transactions from {startDate} to {endDate}"

        'req.Content = New FormUrlEncodedContent(d)
        Dim parms = New FormUrlEncodedContent(d)
        'Dim query As String = "" ' parms.ReadAsStringAsync().Result

        ' https://stackoverflow.com/questions/56823997/add-parameters-to-httpclient
        ' https://stackoverflow.com/questions/10679214/how-to-set-the-content-type-header-for-an-httpclient-request
        ' https://stackoverflow.com/questions/78197923/how-to-add-parameters-when-using-httpclient-to-send-request-net-core-c
        ' https://stackoverflow.com/questions/20532711/posting-with-c-sharp-httpclient-with-formencoded-paramaters-And-headers

        Dim queryString = New StringBuilder()
        Dim delim As String = "?"
        For Each k As String In d.Keys
            queryString.Append($"{delim}{k}=").Append(Uri.EscapeDataString(d(k)))
            delim = "&"
        Next
        'queryString.Append("?UN=").Append(Uri.EscapeDataString("aa"))
        'queryString.Append("&AP=").Append(Uri.EscapeDataString("bb"))



        Dim req = New HttpRequestMessage(HttpMethod.Get, "https://login.microsoftonline.com/0475dfa7-xxxxxxxx-896cf5e31efc/oauth2/token")
        req.Headers.Add("Referer", "login.microsoftonline.com")
        req.Headers.Add("Accept", "application/x-www-form-urlencoded")
        'req.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
        'req.Content = parms


        ' Dim res As HttpResponseMessage = Await clientJPMC.GetAsync("transactions")
        ' Dim res As HttpResponseMessage = Await clientJPMC.GetAsync("transactions")
        ' Dim res2 As HttpResponseMessage = Await clientJPMC.GetAsync("transactions?" & query)
        ' Dim res2 As HttpResponseMessage = Await clientJPMC.SendAsync(req)

        Dim reqURL = "transactions"
        reqURL &= queryString.ToString
        Dim res2 As HttpResponseMessage = Await clientJPMC.GetAsync(reqURL)
        'Dim res2 As HttpResponseMessage = Await clientJPMC.GetAsync("transactions")
        'Dim res2 As HttpResponseMessage = Await clientJPMC.GetAsync("transactions")

        'Dim resp As HttpResponseMessage = Await clientJPMC.SendAsync(req)

        'https://openbankinguat.jpmorgan.com/tsapi/v3/transactions?pageNumber=1&accountIds=&relativeDateType=PRIOR_DAY

        ' https://stackoverflow.com/questions/17096201/build-query-string-for-system-net-httpclient-get

        '        String query;
        'Using (var content = New FormUrlEncodedContent(New KeyValuePair<string, string>[]{
        '    New KeyValuePair<string, string>("ham", "Glazed?"),
        '    New KeyValuePair<string, string>("x-men", "Wolverine + Logan"),
        '    New KeyValuePair<string, string>("Time", DateTime.UtcNow.ToString()),
        '})) {
        '    query = content.ReadAsStringAsync().Result;
        '}




        Dim body As String = Await res2.Content.ReadAsStringAsync()
        Dim c As JPMC_Transactions_Response = JsonConvert.DeserializeObject(Of JPMC_Transactions_Response)(body)

        txtTransactions.Text = body

        Dim cc As New List(Of JPMC_Transactions_Response)
        cc.Add(c)
        grdTransactions.DataSource = cc


    End Sub

    Private Sub btnData_Click(sender As Object, e As EventArgs) Handles btnData.Click

        Dim tbl As New DataTable
        With tbl
            .Columns.Add("KEY")
            .Columns.Add("DESC")
            .Columns.Add("STATUS")
            .Columns.Add("ACTIVE")
            .Columns("ACTIVE").DefaultValue = "0"
        End With


        tbl.Rows.Add(New String() {"1", "Record 1", "A", "1"})
        tbl.Rows.Add(New String() {"2", "Record 2", "I", "0"})

        grdData.DataSource = tbl

        Dim VL As New ValueList
        VL.ValueListItems.Add(New ValueListItem("A", "Active"))
        VL.ValueListItems.Add(New ValueListItem("I", "Inactive"))
        grdData.DisplayLayout.Bands(0).Columns("STATUS").ValueList = VL

        grdData.DisplayLayout.Bands(0).Columns("ACTIVE").Style = UltraWinGrid.ColumnStyle.CheckBox
        ASCMAIN1.grdInitializeLayout(grdData)
    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Dim tbl As DataTable = DirectCast(grdData.DataSource, DataTable)
        Dim Data As String = ""
        For Each row As DataRow In tbl.Rows
            Data &= Join(row.ItemArray, " : ") & vbCrLf
        Next

        MsgBox("Data" & vbCrLf & Data)
    End Sub
End Class

Public Class Account
    Public Property accountId As String
    Public Property accountName As String
    Public Property bankId As String
    Public Property branchId As String
    Public Property bankName As String
    Public Property aba As String
    Public Property swift As Object
    Public Property currency As Currency
End Class

Public Class BaiType
    Public Property typeCode As String
    Public Property description As String
    Public Property btrsTypeCode As String
End Class

Public Class BankReferenceSearchable
    Public Property standardValue As String
End Class

Public Class Currency
    Public Property code As String
    Public Property description As String
End Class

Public Class CustomerReferenceSearchable
    Public Property standardValue As String
End Class

Public Class Datum
    Public Property account As Account
    Public Property asOfDateTime As DateTime
    Public Property valueDateTime As DateTime
    Public Property asOfDate As String
    Public Property valueDate As String
    Public Property receivedTimestamp As DateTime
    Public Property debitCreditCode As String
    Public Property baiType As BaiType
    Public Property fundsTypeCode As String
    Public Property currency As Currency
    Public Property amount As Double
    Public Property immediateAvailable As Double
    Public Property day1Available As Double
    Public Property day2Available As Double
    Public Property day2PlusAvailable As Object
    Public Property day3PlusAvailable As Double
    Public Property bankReferenceSearchable As BankReferenceSearchable
    Public Property customerReferenceSearchable As CustomerReferenceSearchable
    Public Property repairCode As String
    Public Property reversal As Boolean
    Public Property checkNumber As Integer
    Public Property wireType As String
    Public Property shortDescription As String
    Public Property postCode As String
    Public Property lockbox As Lockbox
    Public Property narrativeText As NarrativeText
    Public Property addenda As List(Of Object)
    Public Property sepaDetailsXml As Object
    Public Property supplementalTextSet As SupplementalTextSet
    Public Property supplementalTextRecordList As Object
    Public Property supplementalText As Object
    Public Property achBatchItems As Object
    Public Property transactionId As String
End Class

Public Class Lockbox
    Public Property lockboxSequenceCode As String
    Public Property lockboxItems As Double
    Public Property lockboxNumber As String
    Public Property lockboxDepositDate As Object
    Public Property lockboxDepositTime As Object
End Class

Public Class NarrativeText
    <JsonProperty("YOUR REF    ")>
    Public Property YOURREF As String
    <JsonProperty("REC FROM    ")>
    Public Property RECFROM As String
    <JsonProperty("REMARK      ")>
    Public Property REMARK As String
    <JsonProperty("REC GFP     ")>
    Public Property RECGFP As String
    <JsonProperty("B/O CUSTOMER")>
    Public Property BOCUSTOMER As String
    <JsonProperty("B/O BANK    ")>
    Public Property BOBANK As String
    <JsonProperty("CHIP SEQ    ")>
    Public Property CHIPSEQ As String
    <JsonProperty("CHIP REF    ")>
    Public Property CHIPREF As String
    <JsonProperty("ACCT PARTY  ")>
    Public Property ACCTPARTY As String
    <JsonProperty("ULTI BENE   ")>
    Public Property ULTIBENE As String
    <JsonProperty("PAID TO     ")>
    Public Property PAIDTO As String
End Class

Public Class Pagination
    Public Property pageSize As Integer
    Public Property totalPages As Integer
    Public Property pageNumber As Integer
    Public Property totalRecords As Integer
End Class

Public Class JPMC_Transactions_Response
    Public Property pagination As Pagination
    Public Property data As List(Of Datum)
End Class

Public Class SupplementalTextSet
End Class

'Public Class JWTSettings
'    ' Implements IJWTSettings

'    Public Sub New()
'    End Sub

'    Public Property Key As String
'    Public Property Issuer As String
'    Public Property Audience As String
'    Public Property Subject As String
'    Public Property TimeOfIssuance As Int64
'    Public Property TimeOfExpiration As Int64
'    Public Property UniqueIdentifier As Int64
'    Public Property Thumbprint As String
'    Public Property TokenType As String
'    Public Property EncryptionAlgorithm As String

'End Class

Public Class person
    Public Property Name As String
    Public Property Email As String
    Public Property Address As String
End Class
