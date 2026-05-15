Public Class SAR12MO2
    Dim RYPs(1, 12) As String
    Dim LEGENDs(1, 12) As String

    Dim PP As Integer = 0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim RYP01 As String = Mid(RYP, 1, 4) & "01"

        ' Prepare Work Tables

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        Dim sql_Cols_Retail As String = ""

        ASCMAIN1.Progress("Shipments")

        MyBase.Get_SQL("*")

        sql_Data = ""
        sql_Cols = ""
        For Each YY As String In New String() {"TY", "LY"}
            Dim Y As Integer = (IIf(YY = "TY", 0, 1))
            For M As Integer = 1 To 12
                Dim YP As String = ""
                If YY = "TY" Then
                    YP = Mid(RYP, 1, 4) & Format(M, "00")   
                Else
                    YP = ASCMAIN1.Period_Calc(RYPs(0, M), -12)
                End If

                sql_Cols_Retail &= ", MAX(DECODE(OPS_YYYYPP,'" & YP & "',ITEM_RETAIL_PRICE,0)) " & YY & "_SRP_" & Format(M, "00")

                RYPs(Y, M) = YP
                LEGENDs(Y, M) = ASCMAIN1.Get_Legend(YP)

                If M <= Val(Mid(RYP, 5, 2)) Or Y = 1 Then
                    sql_Data &= "" _
                   & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) " & YY & "_UNITS_" & Format(M, "00") & vbCrLf _
                   & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) " & YY & "_SALES_" & Format(M, "00") & vbCrLf
                Else
                    sql_Data &= "" _
                    & ", 0 " & YY & "_UNITS_" & Format(M, "00") & vbCrLf _
                    & ", 0 " & YY & "_SALES_" & Format(M, "00") & vbCrLf

                End If
                sql_Cols &= "" _
                    & "," & YY & "_UNITS_" & Format(M, "00") & "," & YY & "_SALES_" & Format(M, "00")
            Next M
        Next

        sql_filter = " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & ASCMAIN1.Period_Calc(RYP01, -12) & "' and '" & RYP & "'" & vbCrLf _
            & " and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & " and SOTINVH2.ORDR_QTY_SHIP <> 0"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        Create_TDA(dst.Tables.Add, "ICTRETLX", "Select ITEM_CODE" & sql_Cols_Retail & " from ICTRETLA group by ITEM_CODE", 0, False, "", 1)
        Fill_Records("ICTRETLX")


        Create_TDA(dst.Tables.Add, "ICTITEM1", "Select ITEM_CODE, ITEM_DESC from ICTITEM1", 0, False, "", 1)
        Fill_Records("ICTITEM1")

        ' Create_Pivot()
        ' Load_Excel()


        ' Eliminate 0s

        'Dim sqlz As String = ""
        'For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
        '    sqlz &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        'Next
        'ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlz))
    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
        Load_Excel()
    End Sub

    Sub Load_Excel()

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        Dim XL_ROWS As Integer
        Dim XL_COLS As Integer

        ' Set up parameterized row and col settings

        XL_ROWS = dst.Tables("ASTSRPT1").Rows.Count ' # of Rows in Work Table
        XL_COLS = 0   ' # of numeric columns in Layout Selected

        Dim MX As Integer = 6 + 1 + 6 + 1 + 7 + 1

        Dim PP As Integer = Val(Mid(RYP, 5, 2))

        ' Create Workbook

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        oSheet = oWB.Sheets(0)
        oSheet.Name = "Items"


        ' Load the DataTable into the Item Summary Sheet


        Dim TR As Integer = 3
        Dim TC As Integer = 0

        '  Load_DataTable_into_SGXLS(TR, TC, dst.Tables("ASTSRPT1"), oSheet, Nothing, Nothing, "G1,G2,G3,G4,G5,G6,G7,G8,G9", "")

        Dim Rx As Integer = -1

        Dim c As Integer = TC
        Dim R As Integer = 0

        ' Headings

        R = TR + Rx

        c = TC

        For I As Integer = 1 To COLUMN_CAPTIONs.Count
            oSheet.Cells(R, c + I).Value = COLUMN_CAPTIONs(I - 1)
            oSheet.Cells(R, c + I).EntireColumn.ColumnWidth = 10
        Next
        If COLUMN_CAPTIONs.Count < 9 Then
            For I As Integer = COLUMN_CAPTIONs.Count + 1 To 9
                oSheet.Cells(R, c + I).EntireColumn.Hidden = True
            Next
        End If

        c = TC + 9 + 1
        oSheet.Cells(R, c).Value = "Description"
        oSheet.Cells(R, c).EntireColumn.ColumnWidth = 50
        oSheet.Cells(R, c + 1).EntireColumn.ColumnWidth = 2

        oSheet.Cells(R, c).EntireRow.RowHeight = 40

        range = oSheet.Cells(R, TC + 1, R, TC + 9 + 1)

        range.EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
        range.Interior.Color = SpreadsheetGear.Colors.Lavender
        range.EntireColumn.NumberFormat = "@"

        Dim S(14, 1) As Integer

        For M As Integer = 1 To 14
            Dim MM As String = Format(M, "00")

            For Y As Integer = 0 To 1
                Dim XY As String = IIf(Y = 0, "TY", "LY")

                c = TC + 9 + 2 + 1 + (M - 1) * MX + Y * 7

                S(M, Y) = c

                'oSheet.Cells(R, c + 0).Value = XY & " SRP"
                With oSheet.Cells(R, c + 0)
                    .Value = XY & " SRP"
                    .EntireColumn.NumberFormat = "$#,##0.00"
                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With


                'oSheet.Cells(R, c + 1).Value = XY & " Units"
                With oSheet.Cells(R, c + 1)
                    .Value = XY & " Units"
                    .EntireColumn.NumberFormat = "#,##0"
                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.PaleGreen
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

                '  oSheet.Cells(R, c + 2).Value = XY & " Sales @SRP"
                With oSheet.Cells(R, c + 2)
                    .Value = XY & " Sales @SRP"
                    .EntireColumn.NumberFormat = "$#,##0"

                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.PaleGreen
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

                'oSheet.Cells(R, c + 3).Value = XY & " Act Sales"
                With oSheet.Cells(R, c + 3)
                    .Value = XY & " Act Sales"
                    .EntireColumn.NumberFormat = "$#,##0"

                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

                'oSheet.Cells(R, c + 4).Value = XY & " Avg SP"
                With oSheet.Cells(R, c + 4)
                    .Value = XY & " Avg SP"
                    .EntireColumn.NumberFormat = "$#,##0.00"
                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

                ' oSheet.Cells(R, c + 5).Value = XY & " Avg Disc%"
                With oSheet.Cells(R, c + 5)
                    .Value = XY & " Avg Disc%"
                    .EntireColumn.NumberFormat = "#,##0.0%"
                    .EntireColumn.ColumnWidth = 9
                    .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With


                If M = 13 Then
                    oSheet.Cells(R - 1, c + 0).Value = "Year to Date " & XY
                ElseIf M = 14 Then
                    oSheet.Cells(R - 1, c + 0).Value = "Total Year " & XY
                Else
                    oSheet.Cells(R - 1, c + 0).Value = Mid(LEGENDs(Y, M), 10, 6)
                End If

                range = oSheet.Cells(R, c + 0, R, c + 5)
                range.WrapText = True

                range = oSheet.Cells(R - 1, c + 0, R - 1, c + 5)
                range.Merge()
                If M = 13 Or M = 14 Then oSheet.Cells(R - 1, c + 0).EntireColumn.Hidden = True
                range.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                range.Interior.Color = SpreadsheetGear.Colors.LightBlue

                oSheet.Cells(R, c + 5 + 1).EntireColumn.ColumnWidth = 1
            Next

            c = TC + 9 + 2 + 1 + (M - 1) * MX + 2 * 7

            oSheet.Cells(R - 1, c + 0).Value = "Variances"


            range = oSheet.Cells(R, c + 0, R, c + 6)
            range.WrapText = True

            range = oSheet.Cells(R - 1, c + 0, R - 1, c + 6)
            range.Merge()
            range.HorizontalAlignment = SpreadsheetGear.HAlign.Center
            range.Interior.Color = SpreadsheetGear.Colors.LightGray

            c += 0 ' : oSheet.Cells(R, c).Value = "Var Units"
            With oSheet.Cells(R, c)
                .Value = "Var Units"
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGreen
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1 : oSheet.Cells(R, c).Value = "Var% Units"
            With oSheet.Cells(R, c)
                .Value = "Var% Units"
                .EntireColumn.NumberFormat = "#,##0.0%"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGreen
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1 ' : oSheet.Cells(R, c).Value = "Var Sales @SRP"
            With oSheet.Cells(R, c)
                .Value = "Var Sales @SRP"
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1 ' : oSheet.Cells(R, c).Value = "Var% Sales @SRP"
            With oSheet.Cells(R, c)
                .Value = "Var% Sales @SRP"
                .EntireColumn.NumberFormat = "#,##0.0%"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1 ' : oSheet.Cells(R, c).Value = "Var Act Sales"
            With oSheet.Cells(R, c)
                .Value = "Var Act Sales"
                .EntireColumn.NumberFormat = "#,##0"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1 ' : oSheet.Cells(R, c).Value = "Var% Act Sales"
            With oSheet.Cells(R, c)
                .Value = "Var% Act Sales"
                .EntireColumn.NumberFormat = "#,##0.0%"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.PaleGoldenrod
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            c += 1  ' : oSheet.Cells(R, c).Value = "Var% Disc%"
            With oSheet.Cells(R, c)
                .Value = "Var% Disc%"
                .EntireColumn.NumberFormat = "#,##0.0%"
                .EntireColumn.ColumnWidth = 9
                .Interior.Color = SpreadsheetGear.Colors.LightGray
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With


            c += 1
            oSheet.Cells(R, c).EntireColumn.ColumnWidth = 2

        Next


        ' Data

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
            Rx += 1
            R = TR + Rx

            c = TC

            Dim ITEM_CODE As String = ""

            For I As Integer = 1 To COLUMN_CAPTIONs.Count
                Dim DATA As String = row.Item("G" & CStr(I))
                DATA = Split(DATA & ":", ":")(1)
                oSheet.Cells(R, c + I).Value = DATA
                If I = COLUMN_CAPTIONs.Count Then
                    ITEM_CODE = DATA
                End If
            Next

            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            Dim rowICTRETLX As DataRow = dst.Tables("ICTRETLX").Rows.Find(ITEM_CODE)

            c = TC + 9 + 1
            oSheet.Cells(R, c).Value = rowICTITEM1.Item("ITEM_DESC") & ""

            Dim XY_UNITS_YTD(1) As String
            Dim XY_UNITS_TOT(1) As String
            Dim XY_SALES_YTD(1) As String
            Dim XY_SALES_TOT(1) As String

            For M As Integer = 1 To 14
                Dim MM As String = Format(M, "00")

                For Y As Integer = 0 To 1
                    Dim XY As String = IIf(Y = 0, "TY", "LY")

                    Dim price_change As Boolean = False

                    Dim ITEM_RETAIL_PRICE As Decimal = 0
                    If M <= 12 AndAlso rowICTRETLX IsNot Nothing Then
                        ITEM_RETAIL_PRICE = Val(rowICTRETLX.Item(XY & "_SRP_" & MM) & "")
                    End If
                    If Y = 1 And M <= 12 Then
                        Dim ITEM_RETAIL_PRICE_TY As Decimal = Val(rowICTRETLX.Item("TY" & "_SRP_" & MM) & "")
                        If ITEM_RETAIL_PRICE <> ITEM_RETAIL_PRICE_TY Then
                            price_change = True
                        End If
                    End If

                    c = TC + 9 + 2 + 1 + (M - 1) * MX + Y * 7

                    If M = 13 Or M = 14 Then
                        oSheet.Cells(R, c + 1).Value = Replace(IIf(M = 13, XY_UNITS_YTD(Y), XY_UNITS_TOT(Y)), "+", "=", 1, 1) ' XY Units
                        oSheet.Cells(R, c + 3).Value = Replace(IIf(M = 13, XY_SALES_YTD(Y), XY_SALES_TOT(Y)), "+", "=", 1, 1) ' XY Sales
                    Else
                        oSheet.Cells(R, c + 0).Value = ITEM_RETAIL_PRICE ' XY SRP
                        If price_change Then
                            oSheet.Cells(R, c + 0).Font.Color = SpreadsheetGear.Colors.Red
                        End If

                        Dim UNITS As Int64 = Val(row.Item(XY & "_UNITS_" & MM) & "")
                        oSheet.Cells(R, c + 1).Value = UNITS ' XY Units

                        If M <= PP Then XY_UNITS_YTD(Y) &= "+" & Excel_Cell0(R, c + 1)
                        XY_UNITS_TOT(Y) &= "+" & Excel_Cell0(R, c + 1)

                        oSheet.Cells(R, c + 2).Formula = "=" & Excel_Cell0(R, c + 0) & "*" & Excel_Cell0(R, c + 1) ' XY Gross @SRP
                        oSheet.Cells(R, c + 3).Value = Val(row.Item(XY & "_SALES_" & MM) & "") ' XY Sales

                        If M <= PP Then XY_SALES_YTD(Y) &= "+" & Excel_Cell0(R, c + 3)
                        XY_SALES_TOT(Y) &= "+" & Excel_Cell0(R, c + 3)
                    End If

                    oSheet.Cells(R, c + 4).Formula = "=IF(" & Excel_Cell0(R, c + 1) & "=0,0," & Excel_Cell0(R, c + 3) & "/" & Excel_Cell0(R, c + 1) & ")" ' XY Avg SP
                    oSheet.Cells(R, c + 5).Formula = "=IF(" & Excel_Cell0(R, c + 2) & "=0,0,1-" & Excel_Cell0(R, c + 3) & "/" & Excel_Cell0(R, c + 2) & ")" ' XY Avg Disc%
                Next


                c = TC + 9 + 2 + 1 + (M - 1) * MX + 2 * (6 + 1)

                oSheet.Cells(R, c + 0).Formula = "=" & Excel_Cell0(R, S(M, 0) + 1) & "-" & Excel_Cell0(R, S(M, 1) + 1) ' Var Units
                oSheet.Cells(R, c + 1).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 1) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 1) & "-" & Excel_Cell0(R, S(M, 1) + 1) & ")/" & Excel_Cell0(R, S(M, 1) + 1) & ")" ' Var Units %
                oSheet.Cells(R, c + 2).Formula = "=" & Excel_Cell0(R, S(M, 0) + 2) & "-" & Excel_Cell0(R, S(M, 1) + 2) ' Var Sales @SRP
                oSheet.Cells(R, c + 3).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 2) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 2) & "-" & Excel_Cell0(R, S(M, 1) + 2) & ")/" & Excel_Cell0(R, S(M, 1) + 2) & ")" ' Var Sales @SRP %
                oSheet.Cells(R, c + 4).Formula = "=" & Excel_Cell0(R, S(M, 0) + 3) & "-" & Excel_Cell0(R, S(M, 1) + 3) ' Var Sales
                oSheet.Cells(R, c + 5).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 3) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 3) & "-" & Excel_Cell0(R, S(M, 1) + 3) & ")/" & Excel_Cell0(R, S(M, 1) + 3) & ")" ' Var Sales %
                oSheet.Cells(R, c + 6).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 5) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 5) & "-" & Excel_Cell0(R, S(M, 1) + 5) & ")/" & Excel_Cell0(R, S(M, 1) + 5) & ")" ' Var Disc %
            Next
        Next

        With oSheet.Cells(0, 0)
            .Value = "12 Month Sales Analysis by Item"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = SpreadsheetGear.Colors.Purple
        End With


        ' Border around Entry Area

        Dim R_LAST As Integer = R
        R = 0

        For M As Integer = 1 To 14
            For Y As Integer = 0 To 1
                c = TC + 9 + 2 + 1 + (M - 1) * MX + Y * 7
                With oSheet.Range(TR - 2, c, R_LAST, c + 5)
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).Color = SpreadsheetGear.Colors.LightGray
                    .Borders(SpreadsheetGear.BordersIndex.InsideVertical).Color = SpreadsheetGear.Colors.LightGray
                    '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With


                oSheet.Cells(R, c + 1).Formula = "=SUBTOTAL(9," & Excel_Cell0(TR, c + 1) & ":" & Excel_Cell0(R_LAST, c + 1) & ")"
                oSheet.Cells(R, c + 2).Formula = "=SUBTOTAL(9," & Excel_Cell0(TR, c + 2) & ":" & Excel_Cell0(R_LAST, c + 2) & ")"
                oSheet.Cells(R, c + 3).Formula = "=SUBTOTAL(9," & Excel_Cell0(TR, c + 3) & ":" & Excel_Cell0(R_LAST, c + 3) & ")"

                oSheet.Cells(R, c + 4).Formula = "=IF(" & Excel_Cell0(R, c + 1) & "=0,0," & Excel_Cell0(R, c + 3) & "/" & Excel_Cell0(R, c + 1) & ")" ' XY Avg SP
                oSheet.Cells(R, c + 5).Formula = "=IF(" & Excel_Cell0(R, c + 2) & "=0,0,1-" & Excel_Cell0(R, c + 3) & "/" & Excel_Cell0(R, c + 2) & ")" ' XY Avg Disc%


            Next

            c = TC + 9 + 2 + 1 + (M - 1) * MX + 2 * 7
            With oSheet.Range(TR - 2, c, R_LAST, c + 6)
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).Weight = SpreadsheetGear.BorderWeight.Thin
                '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).Weight = SpreadsheetGear.BorderWeight.Thin
                .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).Color = SpreadsheetGear.Colors.LightGray
                .Borders(SpreadsheetGear.BordersIndex.InsideVertical).Color = SpreadsheetGear.Colors.LightGray
                '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Dash
                '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Dash
            End With

            oSheet.Cells(R, c + 0).Formula = "=" & Excel_Cell0(R, S(M, 0) + 1) & "-" & Excel_Cell0(R, S(M, 1) + 1) ' Var Units
            oSheet.Cells(R, c + 1).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 1) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 1) & "-" & Excel_Cell0(R, S(M, 1) + 1) & ")/" & Excel_Cell0(R, S(M, 1) + 1) & ")" ' Var Units %
            oSheet.Cells(R, c + 2).Formula = "=" & Excel_Cell0(R, S(M, 0) + 2) & "-" & Excel_Cell0(R, S(M, 1) + 2) ' Var Sales @SRP
            oSheet.Cells(R, c + 3).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 2) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 2) & "-" & Excel_Cell0(R, S(M, 1) + 2) & ")/" & Excel_Cell0(R, S(M, 1) + 2) & ")" ' Var Sales @SRP %
            oSheet.Cells(R, c + 4).Formula = "=" & Excel_Cell0(R, S(M, 0) + 3) & "-" & Excel_Cell0(R, S(M, 1) + 3) ' Var Sales
            oSheet.Cells(R, c + 5).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 3) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 3) & "-" & Excel_Cell0(R, S(M, 1) + 3) & ")/" & Excel_Cell0(R, S(M, 1) + 3) & ")" ' Var Sales %
            oSheet.Cells(R, c + 6).Formula = "=IF(" & Excel_Cell0(R, S(M, 1) + 5) & "=0,0,(" & Excel_Cell0(R, S(M, 0) + 5) & "-" & Excel_Cell0(R, S(M, 1) + 5) & ")/" & Excel_Cell0(R, S(M, 1) + 5) & ")" ' Var Disc %



            c = TC + 9 + 2 + 1 + (M - 1) * MX + 2 * 7
            With oSheet.Range(TR, c + 6, R_LAST, c + 6)
                .Interior.Color = SpreadsheetGear.Colors.LightGray
            End With

        Next


        ' Column Filters
        c = TC
        oSheet.Cells(R, c + 1, R, c + COLUMN_CAPTIONs.Count).EntireColumn.AutoFilter()
        '   oSheet.Cells(R, c + 9 + 1).EntireColumn.AutoFilter()


        ' Freeze Panes

        oSheet.Range(Excel_Cell0(TR, TC + 1 + COLUMN_CAPTIONs.Count)).Select()
        oSheet.WindowInfo.FreezePanes = True
        oSheet.Range("A1:A1").Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
 

End Class