Public Class DPRABCR1
    Dim ICTITEMX As String = ""
    Dim ABCs As New Dictionary(Of String, Decimal)

    ' these may need to be moved to ASFBASE0 or ASFSRPTM
    Private oWB As SpreadsheetGear.IWorkbook

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_Read_Only(grpInclude, True)
        Set_Read_Only(grpABCGroups, True)


        'If Not licensed Then
        '    Try
        '        SpreadsheetGear.Factory.SetSignedLicense("SpreadsheetGear.License, Type=Standard, Hash=u7vWJ7PJT4WO2DfHsVqHLnA, Product=BND, NewVersionsUntil=2023-02-10, Company=Applied Business Systems'  Inc., Email=wjz@absolution.com, Signature=ZuWI3idCi2Ln0KOz+6nNDbzK3ZLJGD5sKLhIPASp0kg'-#EFsMTU07NhcUYDDvu8cxfaal6j8KRKxjtIzhdSs9y8oA#J")
        '    Catch ex As Exception

        '    End Try
        '    licensed = True
        'End If

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        Do While oWB.Worksheets.Count > 1
            oWB.Worksheets(1).Delete()
        Loop

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sqlw &= SQL_in("ITEM_CATGY_CODE", "ICTITEM1.ITEM_CATGY_CODE")
        sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

        If Absx1.chkFor("CHK_SNU_S").Checked And
           Absx1.chkFor("CHK_SNU_N").Checked And
           Absx1.chkFor("CHK_SNU_U").Checked Then
        Else
            sqlw &= " and ICTITEM1.ITEM_SNU_CODE in (" & Mid(
                 IIf(Absx1.chkFor("CHK_SNU_S").Checked, ",'S'", "") &
                 IIf(Absx1.chkFor("CHK_SNU_N").Checked, ",'N'", "") &
                 IIf(Absx1.chkFor("CHK_SNU_U").Checked, ",'U'", ""), 2) & ")"
        End If

        If Absx1.chkFor("CHK_BP_B").Checked And
           Absx1.chkFor("CHK_BP_P").Checked Then
        Else
            sqlw &= " and ICTITEM1.ITEM_BASIC_PROMO in (" & Mid(
                 IIf(Absx1.chkFor("CHK_BP_B").Checked, ",'B'", "") &
                 IIf(Absx1.chkFor("CHK_BP_P").Checked, ",'P'", ""), 2) & ")"
        End If

        If Absx1.chkFor("CHK_MB_M").Checked And
           Absx1.chkFor("CHK_MB_B").Checked Then
        Else
            sqlw &= " and ICTITEM1.ITEM_BASIC_PROMO in (" & Mid(
                 IIf(Absx1.chkFor("CHK_MB_M").Checked, ",'M'", "") &
                 IIf(Absx1.chkFor("CHK_MB_B").Checked, ",'B'", ""), 2) & ")"
        End If

        RWU = "R"

        Prepare_dst(True, sqlw, RYP)

        Check_if_Empty("ICTITEMX")
    End Sub

    Overrides Function Prepare_dst(
      ByVal perform_fill As Boolean,
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        If ICTITEMX = "" Then Create_Temp_Data(sqlw)

        With dst
            ASCMAIN1.sql = "Select * from " & ICTITEMX
            Create_TDA(dst.Tables.Add("ICTITEMX"), ICTITEMX, "**", 0, True, , 1)
            With .Tables("ICTITEMX").Columns
                .Add("DEMAND_QTY", GetType(System.Int64), "ISNULL(FORECAST,0)+ISNULL(PROD_COM,0)+ISNULL(PLAN_COM,0)")
                .Add("DEMAND_AMT", GetType(System.Decimal), "DEMAND_QTY * ISNULL(ITEM_COST_STD,0)")
                .Add("DEMAND_PCT", GetType(System.Decimal))
                .Add("DEMAND_PCT_CUM", GetType(System.Decimal))
            End With

            Create_TDA(dst.Tables.Add, "ICTCOLL1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTBRAN1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTCATG1", "*", 0)
            Create_TDA(dst.Tables.Add, "DPTABCP1", "*", 0)
            .Tables("DPTABCP1").Columns.Add("ABC_INDEX", GetType(System.Int64))
            .Tables("DPTABCP1").Columns.Add("ABC_PCT_CUM", GetType(System.Decimal))

            With .Tables.Add("DPTABCPG")
                .Columns.Add("ABC_GROUP")
                .Columns.Add("ABC_GROUP_DESC")
                .PrimaryKey = New DataColumn() { .Columns("ABC_GROUP")}
            End With

        End With

        Fill_Records("ICTCOLL1")
        Fill_Records("ICTBRAN1")

        Fill_Records("ICTCATG1")
        dst.Tables("ICTCATG1").Rows.Add(New String() {"*", "All Catgys"})
        dst.Tables("ICTCATG1").Rows.Add(New String() {"?", "Catgy Unknown"})

        Fill_Records("DPTABCP1")

        Dim ABC_INDEX As Int16 = 0
        Dim ABC_PCT_CUM As Decimal
        For Each rowDPTABCP1 As DataRow In dst.Tables("DPTABCP1").Select("", "ABC_CODE")
            ABC_INDEX += 1
            ABC_PCT_CUM += Val(rowDPTABCP1.Item("ABC_PCT_RANGE") & "")
            rowDPTABCP1.Item("ABC_INDEX") = ABC_INDEX
            rowDPTABCP1.Item("ABC_PCT_CUM") = ABC_PCT_CUM
        Next

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)

        If sqlw <> "" Then
            Create_Temp_Data(sqlw)
        End If
        EnforceConstraints(False)
        Fill_Records("ICTITEMX")
        EnforceConstraints(True)

        Calculate_ABC()

    End Sub

    Sub Create_Temp_Data(SQLW As String)

        Dim TBC As String = Absx1.optFor("OPTTBC").Value

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTITEM1.ITEM_UOM, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_PO_QTY_MIN" & vbCrLf _
            & ", ICTITEM1.ITEM_MRP_PLANR_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE || ICTITEM1.ITEM_BASIC_PROMO || ICTITEM1.ITEM_COST_MAKE_BUY || ICTITEM1.ITEM_CATGY_CODE ABC_GROUP" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_POS_MAX, ICTITEM1.ITEM_POS_MIN, ICTITEM1.ITEM_MIN_DAYS_SUPPLY, ICTITEM1.ITEM_ABC_PARMS_LOCKED" & vbCrLf _
            & $", {IIf(TBC = "C", "ICTITEM1.COLLECTION_CODE", "'*' COLLECTION_CODE")}" & vbCrLf _
            & $", {IIf(TBC = "B" Or TBC = "C", "ICTCOLL1.BRAND_CODE", "'*' BRAND_CODE")}" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_MAKE_BUY, ICTITEM1.ITEM_ABC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_ABC_CODE ITEM_ABC_CODE_FUT" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & SQLW

        If ICTITEMX = "" Then
            ICTITEMX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add FORECAST NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add PROD_COM NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEMX & " Add PLAN_COM NUMBER (8,0)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEMX)

            Dim COLUMN_NAMEs As String = "" _
                & "ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ITEM_COST_STD" _
                & ", ITEM_UOM, VEND_CODE, ITEM_PO_QTY_MIN, ITEM_MRP_PLANR_CODE, ABC_GROUP" _
                & ", ITEM_CATGY_CODE, ITEM_POS_MAX, ITEM_POS_MIN, ITEM_MIN_DAYS_SUPPLY, ITEM_ABC_PARMS_LOCKED" _
                & ", COLLECTION_CODE, BRAND_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" _
                & ", ITEM_COST_MAKE_BUY, ITEM_ABC_CODE, ITEM_ABC_CODE_FUT"

            ASCDATA1.ExecuteSQL("Insert into " & ICTITEMX & " (" & COLUMN_NAMEs & ") " & ASCMAIN1.sql)

            Get_Demand_Data()
        End If

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        CR_params.Add("MOS", Val(Absx1.numFor("FPDMOS").Value & ""))
        CR_params.Add("CD", "Demand Calculated using " & Absx1.optFor("OPTDEMAND").Text)
        CR_params.Add("EXTUSAGE", "Saleable Ranked by " & Absx1.optFor("OPTRANKBY").Text)
        Generate_Report(RPT, , SUBT)

        Generate_XLS()
    End Sub

    Sub Generate_XLS()
        'ASCMAIN1.sql = ""
        'Dim tbl0 As DataTable = ASCDATA1.GetDataTable

        Dim tbl0 As DataTable = dst.Tables("ICTITEMX").Copy

        'tbl0.Columns.Add("ITEM_ABC_CODE")

        Dim MSGs As New List(Of String)
        Dim CAs As New Dictionary(Of String, ColumnAttributes)

        MSGs.Clear()
        MSGs.Add($"ABC Classifcation Report")

        CAs.Clear()
        Dim totals As New List(Of String)

        CAs.Add("ITEM_CODE", New ColumnAttributes With {.Caption = "Item Code"})
        CAs.Add("ITEM_DESC", New ColumnAttributes With {.Caption = "Item Description", .Width = 40, .HAlign = SpreadsheetGear.HAlign.Left})
        CAs.Add("ITEM_RETAIL_PRICE", New ColumnAttributes With {.Caption = "Retail Price", .Format = "#,##0.00"})
        CAs.Add("ITEM_COST_STD", New ColumnAttributes With {.Caption = "Std Cost", .Format = "#,##0.0000"})
        CAs.Add("ITEM_UOM", New ColumnAttributes With {.Caption = "UM"})
        CAs.Add("VEND_CODE", New ColumnAttributes With {.Caption = "Vendor"})
        CAs.Add("ITEM_PO_QTY_MIN", New ColumnAttributes With {.Caption = "PO Min Qty", .Format = "#,##0"})
        CAs.Add("ITEM_MRP_PLANR_CODE", New ColumnAttributes With {.Caption = "Planner"})
        CAs.Add("ABC_GROUP", New ColumnAttributes With {.Caption = "ABC Group"})
        CAs.Add("ITEM_CATGY_CODE", New ColumnAttributes With {.Caption = "Catgy"})
        CAs.Add("ITEM_POS_MAX", New ColumnAttributes With {.Caption = "Max Pos", .Format = "#,##0.00"})
        CAs.Add("ITEM_POS_MIN", New ColumnAttributes With {.Caption = "Min Pos", .Format = "#,##0.00"})
        CAs.Add("ITEM_MIN_DAYS_SUPPLY", New ColumnAttributes With {.Caption = "Min Days", .Format = "#,##0"})
        CAs.Add("ITEM_ABC_PARMS_LOCKED", New ColumnAttributes With {.Caption = "Locked"})
        CAs.Add("COLLECTION_CODE", New ColumnAttributes With {.Caption = "Collection"})
        CAs.Add("BRAND_CODE", New ColumnAttributes With {.Caption = "Brand"})
        CAs.Add("ITEM_SNU_CODE", New ColumnAttributes With {.Caption = "SNU"})
        CAs.Add("ITEM_BASIC_PROMO", New ColumnAttributes With {.Caption = "BP"})
        CAs.Add("ITEM_COST_MAKE_BUY", New ColumnAttributes With {.Caption = "MB"})
        CAs.Add("ITEM_ABC_CODE", New ColumnAttributes With {.Caption = "ABC Cur"})
        CAs.Add("ITEM_ABC_CODE_FUT", New ColumnAttributes With {.Caption = "ABC Fut"})

        CAs.Add("FORECAST", New ColumnAttributes With {.Caption = "Forecast", .Format = "#,##0"}) : totals.Add("FORECAST")
        CAs.Add("PROD_COM", New ColumnAttributes With {.Caption = "Prod Comm", .Format = "#,##0"}) : totals.Add("PROD_COM")
        CAs.Add("PLAN_COM", New ColumnAttributes With {.Caption = "Plan Comm", .Format = "#,##0"}) : totals.Add("PLAN_COM")
        CAs.Add("DEMAND_QTY", New ColumnAttributes With {.Caption = "Demand", .Format = "#,##0"}) : totals.Add("DEMAND_QTY")
        CAs.Add("DEMAND_AMT", New ColumnAttributes With {.Caption = "Demand Amt", .Format = "#,##0.00"}) : totals.Add("DEMAND_AMT")
        CAs.Add("DEMAND_PCT", New ColumnAttributes With {.Caption = "Demand Pct", .Format = "#,##0.00"})
        CAs.Add("DEMAND_PCT_CUM", New ColumnAttributes With {.Caption = "Demand Pct Cum", .Format = "#,##0.00"}) : totals.Add("DEMAND_PCT_CUM")

        'AddSheet(New DataTable() {tbl0}, "ABC Classification", 0, CAs, New String() {"ITEM_PO_QTY_MIN"}, MSGs)
        AddSheet(New DataTable() {tbl0}, "ABC Classification", 0, CAs, totals.ToArray, MSGs)

        Dim FILENAME As String = SaveWorkbook_Report(XNO, MENU_ITEM_OBJECT)

    End Sub

    Function SaveWorkbook_Report(XNO As String, MENU_ITEM_OBJECT As String) As String

        Dim FILENAME As String = SaveWorkbook(XNO, MENU_ITEM_OBJECT)
        Show_Document(FILENAME)
        'ASCMAIN1.sql = "Update WBTRPTH1 Set STATUS = 'C', STATUS_DATE_C = SYSDATE where RPT_EXE_NO = :PARM1"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {RPT_EXE_NO})

        Return FILENAME
    End Function

    Function SaveWorkbook(XNO As String, MENU_ITEM_OBJECT As String) As String

        oWB.Worksheets(0).Select()

        Dim reportsHome As String = ASCMAIN1.Folders("Archive") & "XLS\"
        Dim FOLDER As String = System.IO.Path.Combine(reportsHome, MENU_ITEM_OBJECT)

        If Not System.IO.Directory.Exists(FOLDER) Then
            System.IO.Directory.CreateDirectory(FOLDER)
        End If

        oWB.SaveAs(System.IO.Path.Combine(FOLDER, XNO & ".XLSX"), SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Return System.IO.Path.Combine(FOLDER, XNO & ".XLSX")
    End Function

    Public Sub AddSheet(
        tbls() As DataTable,
        SHEET_NAME As String,
        wsIndex As Integer,
        CAs As Dictionary(Of String, ColumnAttributes),
        Optional Totals() As String = Nothing,
        Optional MSGs As List(Of String) = Nothing,
        Optional options As List(Of String) = Nothing)

        If options Is Nothing Then options = New List(Of String)

        Dim r0 As Integer = 5
        Dim r0_Orig As Integer = r0
        Dim c0 As Integer = 2

        Dim TotalsOnBottom As Boolean = options.Contains("TotalsOnBottom")

        If MSGs IsNot Nothing Then
            r0 += MSGs.Count + 1
        End If

        Dim oWS As SpreadsheetGear.IWorksheet
        If wsIndex = 0 Then
            oWS = oWB.Worksheets(0)
        Else
            oWS = oWB.Worksheets.Add
        End If

        Dim COLs As New Dictionary(Of String, String)

        For Each tbl As DataTable In tbls

            ' Format Columns

            Dim c As Integer = -1
            For Each CAkey As String In CAs.Keys
                Dim CA As ColumnAttributes = CAs(CAkey)
                c += 1
                Dim fmt As String = "@"
                Dim hdg As String = ""

                If COLs.ContainsKey(CAkey) Then
                    'COLs.Add(CAkey, GetExcelColumnName(c0 + c + 1))
                Else
                    COLs.Add(CAkey, GetExcelColumnName(c0 + c + 1))
                End If

                'Dim halign As SpreadsheetGear.HAlign = SpreadsheetGear.HAlign.Left
                Dim halign As SpreadsheetGear.HAlign = SpreadsheetGear.HAlign.Center

                If CA.Formula = "" Then
                    Dim dcol As DataColumn = tbl.Columns(CAkey)
                    If dcol IsNot Nothing Then
                        hdg = dcol.Caption

                        Select Case dcol.DataType.ToString
                            Case "System.Decimal", "System.Double"
                                fmt = "#,##0.00"
                                halign = SpreadsheetGear.HAlign.Right
                            Case "System.Integer", "System.Int32", "System.Int64"
                                fmt = "#,##0"
                                halign = SpreadsheetGear.HAlign.Right
                            Case "System.DateTime", "System.Date"
                                fmt = "MM/dd/yyyy"
                                halign = SpreadsheetGear.HAlign.Center
                        End Select
                    End If


                End If

                If CA.Format <> "" Then fmt = CA.Format
                If CA.Caption <> "" Then hdg = CA.Caption
                If Not CA.HAlign.Equals(SpreadsheetGear.HAlign.General) Then
                    halign = CA.HAlign
                End If

                With oWS.Cells(0, c0 + c).EntireColumn
                    .NumberFormat = fmt
                    .HorizontalAlignment = halign
                End With

                With oWS.Cells(r0, c0 + c)
                    .Value = hdg
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .WrapText = True
                End With
            Next

            ' Style the Column Headings area
            With oWS.Cells(r0, c0, r0, c0 + c)
                .Interior.Color = SpreadsheetGear.Colors.LightBlue
                .Font.Color = SpreadsheetGear.Colors.Navy
            End With

            ' Resolve Formula Columns
            For Each CAkey As String In CAs.Keys
                Dim CA As ColumnAttributes = CAs(CAkey)
                Dim F As String = CA.Formula
                If F <> "" Then
                    For Each CAkey2 As String In CAs.Keys
                        If F.IndexOf(CAkey2 & "_total") >= 0 Then
                            F = F.Replace(CAkey2 & "_total", COLs(CAkey2) & "{999999}")
                        End If
                        If F.IndexOf(CAkey2) >= 0 Then
                            F = F.Replace(CAkey2, COLs(CAkey2) & "{000000}")
                        End If
                    Next
                    CA.FormulaResolved = F
                End If
            Next

            ' Determine the row for Totals, and insert a row if we are doing Totals on the Top
            Dim RC As Integer = tbl.Rows.Count
            Dim rT As Integer = r0 + RC + 1
            Dim rX As Integer = 0
            If tbls.Length = 1 And Not TotalsOnBottom Then
                oWS.Cells(r0, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                rT = r0
                rX = 1
            End If

            ' Insert data from DataTable & Formulas
            ' - note that CopyFromDataTable does not work in .Net Std,
            ' - and we don't want it because we now rely on CAs for column placement & formulas
            ' oWS.Range(r0, c0).CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.None)
            Dim iRow As Integer = 0 + rX ' 1
            For Each row As DataRow In tbl.Select("")
                iRow += 1
                Dim iCol As Integer = -1
                For Each CAkey As String In CAs.Keys
                    Dim CA As ColumnAttributes = CAs(CAkey)
                    iCol += 1
                    If CA.Formula = "" Then
                        oWS.Cells(r0 + iRow, c0 + iCol).Value = row.Item(CAkey)
                    Else
                        Dim F As String = CA.FormulaResolved
                        F = F.Replace("{000000}", CStr(r0 + iRow + 1))
                        F = F.Replace("{999999}", CStr(rT + 1))
                        oWS.Cells(r0 + iRow, c0 + iCol).Formula = F
                    End If
                Next
            Next

            ' Add the Column Filter if we have only 1 table and if we are not doing TotalsOnBottom
            'If tbls.Length = 1 And Not TotalsOnBottom Then oWS.Range(r0 + rX, c0 + 0, r0 + rX, c0 + c).AutoFilter()
            If tbls.Length = 1 And Not TotalsOnBottom Then oWS.Range(r0, c0 + 0, r0, c0 + c).AutoFilter()

            '' Style the Column Headings area
            'With oWS.Cells(r0 + 1, c0, r0 + 1, c0 + c)
            '    .Interior.Color = SpreadsheetGear.Colors.LightBlue
            '    .Font.Color = SpreadsheetGear.Colors.Navy
            'End With

            ' Set the Freeze Panes
            oWS.Cells(r0 + rX + 1, 2).Select()
            If tbls.Length = 1 Then oWS.WindowInfo.FreezePanes() = True

            ' Put a light grid around all of the Data Cells
            If RC > 0 Then
                With oWS.Cells(r0 + rX + 1, c0 + 0, r0 + rX + RC, c0 + c)
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders.Color = SpreadsheetGear.Colors.LightSteelBlue
                End With
            End If

            ' Autofit

            c = -1
            For Each CAkey As String In CAs.Keys
                Dim CA As ColumnAttributes = CAs(CAkey)
                c += 1

                With oWS.Cells(r0, c0 + c)
                    If CA.Width <> 0 Then
                        .EntireColumn.ColumnWidth = CA.Width
                    Else
                        .EntireColumn.AutoFit()
                        .EntireColumn.ColumnWidth = oWS.Range(r0, c0 + c).EntireColumn.ColumnWidth * 1.5
                    End If
                End With

            Next

            ' Totals (Top or Bottom)
            If Totals IsNot Nothing AndAlso Totals.Length > 0 Then
                'If tbls.Length = 1 And Not TotalsOnBottom Then
                '    oWS.Cells(r0, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                'End If
                If RC > 0 Then
                    c = -1
                    For Each CAkey As String In CAs.Keys
                        Dim CA As ColumnAttributes = CAs(CAkey)
                        c += 1

                        If Totals.Contains(CAkey) Then
                            Dim CELL1 As String = oWS.Cells(r0 + 1 + rX, c0 + c).Address
                            CELL1.Replace("$", "")
                            Dim CELL2 As String = oWS.Cells(r0 + RC + rX, c0 + c).Address
                            CELL2.Replace("$", "")
                            If CA.Formula = "" Then
                                If tbls.Length = 1 And Not TotalsOnBottom Then
                                    oWS.Cells(rT, c0 + c).Formula = $"=SUBTOTAL(9,{CELL1}:{CELL2})"
                                Else
                                    oWS.Cells(rT, c0 + c).Formula = $"=SUM({CELL1}:{CELL2})"
                                End If
                            Else
                                Dim F As String = CA.FormulaResolved
                                F = F.Replace("{999999}", CStr(rT + 1))
                                F = F.Replace("{000000}", CStr(rT + 1))
                                oWS.Cells(rT, c0 + c).Formula = F
                            End If
                        End If
                    Next
                End If
                oWS.Cells(rT, c0).Value = "Totals"
                With oWS.Cells(rT, c0, rT, c0 + c)
                    .Interior.Color = SpreadsheetGear.Colors.LightGray
                    .Font.Color = SpreadsheetGear.Colors.Navy
                End With

                r0 += 1
            End If

            ' Row Spacer in case there are multiple tables
            r0 += RC + 3
        Next

        ' Logo
        Dim imageFile As String = ASCMAIN1.CLIENT & ".png"
        Dim IMAGE_PATH As String = ASCMAIN1.Folders("Images")
        imageFile = IMAGE_PATH & "ABS\" & imageFile

        Dim width As Double = 0
        Dim height As Double = 0
        Dim image As System.Drawing.Image = Nothing

        Try
            image = System.Drawing.Image.FromFile(imageFile)
            width = image.Width * 14.0 / image.HorizontalResolution
            height = image.Height * 14.0 / image.VerticalResolution
            image.Dispose()
        Catch ex As Exception

        End Try

        Dim windowInfo As SpreadsheetGear.IWorksheetWindowInfo = oWS.WindowInfo
        Dim left As Double = windowInfo.ColumnToPoints(0)
        Dim top As Double = windowInfo.RowToPoints(1)

        Dim IMG As SpreadsheetGear.Shapes.IShape = oWS.Shapes.AddPicture(imageFile, left, top, width, height)

        ' Style the Jumbotron
        With oWS.Cells(1, 2)
            .Value = SHEET_NAME
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = SpreadsheetGear.Colors.Navy
            .EntireRow.RowHeight = 40
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .IndentLevel = 3
            .Interior.Color = SpreadsheetGear.Colors.LightBlue
        End With

        With oWS.Cells(1, 2, 1, 2 + CAs.Count - 1)
            .Merge()
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        With oWS.Cells(r0_Orig - 2, 0)
            .Value = ASCMAIN1.CLIENT '  PARTNER_CODE
            .Font.Bold = True
            .Font.Color = SpreadsheetGear.Colors.Navy
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        With oWS.Cells(r0_Orig - 1, 1)
            .Value = XNO
            .Font.Bold = True
            .Font.Color = SpreadsheetGear.Colors.Navy
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        With oWS.Cells(r0_Orig - 2, 1)
            .Value = $"Report Produced {System.DateTime.Now.ToString("MM/dd/yyyy HH:mm tt")}"
            .Font.Bold = True
        End With

        If MSGs IsNot Nothing Then
            For m As Integer = 0 To MSGs.Count - 1
                With oWS.Cells(r0_Orig + m, 1)
                    .Value = MSGs(m)
                    .Font.Bold = True
                End With
            Next
        End If

        oWS.WindowInfo.DisplayGridlines = False

        oWS.Name = SHEET_NAME
        oWS.Cells(0, 0).Activate()
    End Sub

    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = [String].Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Not Absx1.chkFor("CHK_SNU_S").Checked And
               Not Absx1.chkFor("CHK_SNU_N").Checked And
               Not Absx1.chkFor("CHK_SNU_U").Checked Then
                EMsg &= vbCr & "You must select at least 1 from Saleable/No-Charge/Unfinished"
            End If
            If Not Absx1.chkFor("CHK_BP_B").Checked And
               Not Absx1.chkFor("CHK_BP_P").Checked Then
                EMsg &= vbCr & "You must select at least 1 from Basic/Promo"
            End If
            If Not Absx1.chkFor("CHK_MB_M").Checked And
               Not Absx1.chkFor("CHK_MB_B").Checked Then
                EMsg &= vbCr & "You must select at least 1 from Make/Buy"
            End If

            ABCs.Clear()
            Dim ABC_PCT_RANGE_total As Decimal = 0
            ASCMAIN1.sql = "Select * from DPTABCP1"
            For Each rowDPTABCP1 As DataRow In ASCDATA1.GetDataTable.Select("", "ABC_CODE")
                ABCs.Add(rowDPTABCP1.Item("ABC_CODE"), Val(rowDPTABCP1.Item("ABC_PCT_RANGE") & ""))
                ABC_PCT_RANGE_total = ABC_PCT_RANGE_total + Val(rowDPTABCP1.Item("ABC_PCT_RANGE") & "")
            Next
            If ABCs.Count = 0 Then
                EMsg &= vbCr & "No ABC Codes set up"
            ElseIf ABC_PCT_RANGE_total = 0 Then
                EMsg &= vbCr & "ABC Code %s must add up to 100%"
            End If

        End If

    End Sub

    Overrides Sub Update_Record()

        Update_Record_TDA("ICTITEMX", "1=1")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 Is" & vbCrLf _
            & " Select ICTITEMX.ITEM_CODE, ICTITEMX.ITEM_ABC_CODE_FUT" & vbCrLf _
            & "  , ICTITEMX.ITEM_ABC_PARMS_LOCKED" & vbCrLf _
            & "  , DPTABCP1.ABC_MAX_POS, DPTABCP1.ABC_MIN_POS, DPTABCP1.ABC_MIN_DAYS_SUPPLY" & vbCrLf _
            & "  from " & ICTITEMX & " ICTITEMX, DPTABCP1" & vbCrLf _
            & "  where DPTABCP1.ABC_CODE = ICTITEMX.ITEM_ABC_CODE_FUT;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTITEM1 Set ITEM_ABC_CODE = R1.ITEM_ABC_CODE_FUT" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "   If NVL(R1.ITEM_ABC_PARMS_LOCKED,'0') <> '1' Then" & vbCrLf _
            & "    Update ICTITEM1 Set " & vbCrLf _
            & "     ITEM_POS_MAX = R1.ABC_MAX_POS" & vbCrLf _
            & "   , ITEM_POS_MIN = R1.ABC_MIN_POS" & vbCrLf _
            & "   , ITEM_MIN_DAYS_SUPPLY = R1.ABC_MIN_DAYS_SUPPLY" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"

        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Get_Demand_Data()

        dst.Tables("DPTABCPG").Rows.Clear()

        Dim sqlABC_GROUP As String = ""
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_SNU").Checked, "NVL(ITEM_SNU_CODE,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_BP").Checked, "NVL(ITEM_BASIC_PROMO,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_MB").Checked, "NVL(ITEM_COST_MAKE_BUY,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_CATGY").Checked, "NVL(ITEM_CATGY_CODE,'?')", "'*'")
        'If Absx1.chkFor("CHK_GROUP_SNU").Checked Then ABC_GROUP &= " || NVL(ITEM_SNU_CODE,'?')"
        'If Absx1.chkFor("CHK_GROUP_BP").Checked Then ABC_GROUP &= " || NVL(ITEM_BASIC_PROMO,'?')"
        'If Absx1.chkFor("CHK_GROUP_MB").Checked Then ABC_GROUP &= " || NVL(ITEM_COST_MAKE_BUY,'?')"
        'If Absx1.chkFor("CHK_GROUP_CATGY").Checked Then ABC_GROUP &= " || NVL(ITEM_CATGY_CODE,'?')"

        ASCDATA1.ExecuteSQL("Update " & ICTITEMX & " Set ABC_GROUP = " & Mid(sqlABC_GROUP, 5))

        Dim ITEM_SNU_CODEs As New Dictionary(Of String, String)
        ITEM_SNU_CODEs.Add("*", "All SNU")
        ITEM_SNU_CODEs.Add("S", "Saleable")
        ITEM_SNU_CODEs.Add("N", "No-Charge")
        ITEM_SNU_CODEs.Add("U", "Unfinished")
        ITEM_SNU_CODEs.Add("?", "SNU Unknown")

        Dim ITEM_BASIC_PROMOs As New Dictionary(Of String, String)
        ITEM_BASIC_PROMOs.Add("*", "Basic & Promo")
        ITEM_BASIC_PROMOs.Add("B", "Basic")
        ITEM_BASIC_PROMOs.Add("P", "Promo")
        ITEM_BASIC_PROMOs.Add("?", "BP Unknown")

        Dim ITEM_COST_MAKE_BUYs As New Dictionary(Of String, String)
        ITEM_COST_MAKE_BUYs.Add("*", "Make & Buy")
        ITEM_COST_MAKE_BUYs.Add("M", "Make")
        ITEM_COST_MAKE_BUYs.Add("B", "Buy")
        ITEM_COST_MAKE_BUYs.Add("?", "MB Unknown")

        For Each ITEM_SNU_CODE As String In ITEM_SNU_CODEs.Keys
            For Each ITEM_BASIC_PROMO As String In ITEM_BASIC_PROMOs.Keys
                For Each ITEM_COST_MAKE_BUY As String In ITEM_COST_MAKE_BUYs.Keys
                    For Each rowICTCATG1 As DataRow In dst.Tables("ICTCATG1").Select("")
                        Dim ITEM_CATGY_CODE As String = rowICTCATG1.Item("ITEM_CATGY_CODE")
                        Dim ITEM_CATGY_DESC As String = rowICTCATG1.Item("ITEM_CATGY_DESC")
                        Dim ABC_GROUP As String = ITEM_SNU_CODE & ITEM_BASIC_PROMO & ITEM_COST_MAKE_BUY & ITEM_CATGY_CODE
                        Dim ABC_GROUP_DESC As String =
                            ITEM_SNU_CODEs(ITEM_SNU_CODE) & ", " &
                            ITEM_BASIC_PROMOs(ITEM_BASIC_PROMO) & ", " &
                            ITEM_COST_MAKE_BUYs(ITEM_COST_MAKE_BUY) & ", " &
                            ITEM_CATGY_DESC
                        dst.Tables("DPTABCPG").Rows.Add(New String() {ABC_GROUP, ABC_GROUP_DESC})
                    Next
                Next
            Next
        Next


        ASCDATA1.ExecuteSQL("Update " & ICTITEMX & " Set ITEM_ABC_CODE_FUT = NULL")

        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, Val(Absx1.numFor("FPDMOS").Value & "") - 1)

        Select Case Absx1.optFor("OPTDEMAND").Value
            Case "F"
                ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(DPTITMF1.FORECAST,0)) FORECAST from DPTITMF1" _
                    & " where DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                    & "   and DPTITMF1.OPS_YYYYPP_FC <= '" & NYP & "'" & vbCrLf _
                    & " having Sum (NVL(DPTITMF1.FORECAST,0)) <> 0" & vbCrLf _
                    & " group by DPTITMF1.ITEM_CODE"

            Case Else
                Dim sql_filter_history As String = ""
                If Absx1.optFor("OPTDEMAND").Value = "R" Then
                    sql_filter_history &= " where SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.CYP & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                Else
                    sql_filter_history &= " where SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                End If

                ASCMAIN1.sql = "Select SOTINVH2.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) FORECAST from SOTINVH2" _
                    & sql_filter_history & vbCrLf _
                    & " having Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) <> 0" & vbCrLf _
                    & " group by SOTINVH2.ITEM_CODE"
        End Select

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTITEMX & " Set FORECAST = R1.FORECAST where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", NYP)
        Dim PRD_END_DATE As String = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")

        Select Case Absx1.optFor("OPTDEMAND").Value
            Case "F"
                ASCMAIN1.sql = "Select POTORDR9.ITEM_CODE ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(POTORDR9.PO_QTY_COM,0)) PROD_COM from POTORDR9,POTORDR2,DPTPLAN1" _
                    & " where POTORDR2.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO (+) = POTORDR9.PO_ORDER_LNO" & vbCrLf _
                    & "   and DPTPLAN1.PLAN_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
                    & "   and NVL(POTORDR2.PO_DATE_COMPSDUE,DPTPLAN1.DATE_COMPSDUE) <= '" & PRD_END_DATE & "'" & vbCrLf _
                    & " having Sum (NVL(POTORDR9.PO_QTY_COM,0)) <> 0" & vbCrLf _
                    & " group by POTORDR9.ITEM_CODE"

            Case Else
                Dim sql_filter_history As String = ""
                If Absx1.optFor("OPTDEMAND").Value = "R" Then
                    sql_filter_history &= " where ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                    sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                Else
                    sql_filter_history &= " where ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                    sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                End If

                ASCMAIN1.sql = "Select ICTSTAT1.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(ICTSTAT1.WHSE_QTY_CON,0)) PROD_COM from ICTSTAT1" _
                    & sql_filter_history & vbCrLf _
                    & " having Sum (NVL(ICTSTAT1.WHSE_QTY_CON,0)) <> 0" & vbCrLf _
                    & " group by ICTSTAT1.ITEM_CODE"
        End Select

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTITEMX & " Set PROD_COM = R1.PROD_COM where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Calculate_ABC()
        With dst.Tables("ICTITEMX").Columns("DEMAND_AMT")
            If Absx1.optFor("OPTRANKBY").Value = "S" Then
                .Expression = "DEMAND_QTY * ISNULL(ITEM_COST_STD,0)"
            ElseIf Absx1.optFor("OPTRANKBY").Value = "R" Then
                .Expression = "DEMAND_QTY * ISNULL(ITEM_RETAIL_PRICE,0)"
            ElseIf Absx1.optFor("OPTRANKBY").Value = "W" Then
                .Expression = "DEMAND_QTY * ISNULL(ITEM_RETAIL_PRICE,0) * .4"
            End If
        End With

        For Each row As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("ICTITEMX"), New String() {"BRAND_CODE", "COLLECTION_CODE", "ABC_GROUP"}).Rows
            Dim BRAND_CODE As String = row.Item("BRAND_CODE") & ""
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""
            Dim ABC_GROUP As String = row.Item("ABC_GROUP") & ""

            Dim sqlw As String = ""
            sqlw &= " and " & IIf(BRAND_CODE = "", "BRAND_CODE is Null", "BRAND_CODE = '" & BRAND_CODE & "'")
            sqlw &= " and " & IIf(COLLECTION_CODE = "", "COLLECTION_CODE is Null", "COLLECTION_CODE = '" & COLLECTION_CODE & "'")
            sqlw &= " and " & IIf(ABC_GROUP = "", "ABC_GROUP is Null", "ABC_GROUP = '" & ABC_GROUP & "'")
            sqlw = Mid(sqlw, 6)

            Dim DEMAND_AMT_CUM As Decimal = Val(dst.Tables("ICTITEMX").Compute("SUM(DEMAND_AMT)", sqlw) & "")
            Dim DEMAND_PCT_CUM As Decimal = 0

            Dim ABC_INDEX As Integer = 0
            Dim ABC_CODE As String = ""
            Dim rowDPTABCP1 As DataRow = Nothing
            Dim ABC_PCT_CUM As Decimal = 0

            For Each rowICTITEMX As DataRow In dst.Tables("ICTITEMX").Select(sqlw, "DEMAND_AMT DESC")
                Dim DEMAND_AMT As Decimal = Val(rowICTITEMX("DEMAND_AMT") & "")
                Dim DEMAND_PCT As Decimal = 0
                If DEMAND_AMT_CUM <> 0 Then DEMAND_PCT = 100 * DEMAND_AMT / DEMAND_AMT_CUM
                DEMAND_PCT_CUM += DEMAND_PCT
                rowICTITEMX.Item("DEMAND_PCT") = DEMAND_PCT
                rowICTITEMX.Item("DEMAND_PCT_CUM") = DEMAND_PCT_CUM
                If (ABC_INDEX = 0 Or DEMAND_PCT_CUM > ABC_PCT_CUM) And ABC_INDEX < ABCs.Count Then
                    ABC_INDEX += 1
                    ABC_CODE = ABCs.Keys(ABC_INDEX - 1)
                    rowDPTABCP1 = dst.Tables("DPTABCP1").Rows.Find(ABC_CODE)
                    ABC_PCT_CUM = Val(rowDPTABCP1.Item("ABC_PCT_CUM") & "")
                End If
                rowICTITEMX.Item("ITEM_ABC_CODE_FUT") = ABC_CODE
            Next
        Next
    End Sub
End Class


Public Class ColumnAttributes
    Property Key As String
    Property Caption As String
    Property Format As String
    Property Backcolor As SpreadsheetGear.Color
    Property Forecolor As SpreadsheetGear.Color
    Property HAlign As SpreadsheetGear.HAlign
    Property Formula As String
    Property FormulaResolved As String
    Property Width As Integer
    Property SuperCaption As String
    Property SubTotal As Boolean
    Property isSpacer As Boolean
    Property spacerWidth As Integer = 10
End Class