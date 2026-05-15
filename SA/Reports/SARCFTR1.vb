Public Class SARCFTR1
     
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
         Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim MM(1, 12) As String ' 1 = LY 0 = TY,  1-12
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)

        Dim FACTOR As Integer = 1
        'If Absx1.chkFor("THOUSANDS").Checked Then
        '    FACTOR = 1000
        'End If


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        Dim COLUMN_NAME As String = "DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'000000',ORDR_QTY_SHIP * ORDR_UNIT_PRICE,0) / " & CStr(FACTOR)

        sql_filter = "" _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '" & Mid(LYP, 1, 4) & "01" & "' AND '" & RYP & "'" & vbCrLf _
            & " and NVL(SOTINVH2.ORDR_UNIT_PRICE,0) <> 0" & vbCrLf _
            & " and NVL(SOTINVH2.ORDR_QTY_SHIP,0) <> 0"

        Dim sql_Data As String = ""

        For Each Y As String In New String() {"TY", "LY"}
            Dim YP As String = Mid(LYP, 1, 4) & "01"
            If Y = "TY" Then YP = Mid(RYP, 1, 4) & "01"
            For M As Integer = 1 To 12
                Dim XYP As String = ASCMAIN1.Period_Calc(YP, M - 1)
                sql_Data &= ", Sum (" & Replace(COLUMN_NAME, "000000", XYP) & ") " & Y & "_M" & Format(M, "00") & vbCrLf
            Next
        Next

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", Case when SOTINVH2.INV_TYPE = 'C' then '3' else Case When ICTITEM1.ITEM_BASIC_PROMO = 'B' then '1' else '2' End End" & vbCrLf _
        & sql_Data _
        & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols & vbCrLf _
        & ", Case when SOTINVH2.INV_TYPE = 'C' then '3' else Case When ICTITEM1.ITEM_BASIC_PROMO = 'B' then '1' else '2' End End"

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended _
        & ", TY_M01, TY_M02, TY_M03, TY_M04, TY_M05, TY_M06, TY_M07, TY_M08, TY_M09, TY_M10, TY_M11, TY_M12" _
        & ", LY_M01, LY_M02, LY_M03, LY_M04, LY_M05, LY_M06, LY_M07, LY_M08, LY_M09, LY_M10, LY_M11, LY_M12)" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()





        ' DO BUDGETS LATER
        'Dim RSTBUDR1 As String = TAC.RSCMAIN1.RSTBUDR1_as_YP()

        'If Absx1.chkFor("THOUSANDS").Checked Then
        '    ASCMAIN1.sql = "Update " & RSTBUDR1 & " Set BUDGET = BUDGET / 1000"
        '    ASCDATA1.ExecuteSQL()
        'End If

        'MyBase.Get_SQL("B")

        'sql_filter = ""

        'sql_Data = ""
        'For Y As Int16 = 0 To 1
        '    sql_Data &= "" _
        '    & ", 0 " & IIf(Y = 0, "TY", "LY") & "_WTD_B"
        'Next

        'For M As Integer = 1 To 12
        '    sql_Data &= "" _
        '    & ", Sum (CASE WHEN OPS_YYYYPP = '" & MM(0, M) & "' THEN BUDGET ELSE 0 END) " & "TY_B" & Format(M, "00") & vbCrLf
        'Next

        'sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
        '& " from " & RSTBUDR1 & " RSTBUDR1 " & sql_TABLE_NAMEs & vbCrLf _
        '& ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        '& " group by " & sql_GROUP_BY_cols

        'ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        '& "(" & G1thru9 & COLUMN_NAMEs_appended _
        '& ", TY_B01, TY_B02, TY_B03, TY_B04, TY_B05, TY_B06, TY_B07, TY_B08, TY_B09, TY_B10, TY_B11, TY_B12" _
        '& ")" & vbCrLf _
        '& "(" & sql & ")"
        'ASCDATA1.ExecuteSQL()


        Dim sqlx As String = ""
        For Each COLUMN_NAME In COLUMN_NAME_sum.Keys
            sqlx &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlx))
    End Sub
 
    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Sub Prepare_Data_Extracts()

        Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy

        For iRow As Int64 = tbl.Rows.Count - 1 To 0 Step -1 '  Each row As DataRow In tbl.Select("")
            Dim row As DataRow = tbl.Rows(iRow)
            For i As Integer = 1 To tblASTDSQLA.Select("SEQUENCE IS NOT NULL", "SEQUENCE").Length
                Dim C As String = row.Item("G" & CStr(i))
                If C = aRC Then
                    row.Delete()
                    Exit For
                Else
                    row.Item("G" & CStr(i)) = Split(C, ":")(1)
                End If
            Next
        Next

        grdASTEXPT1.DataSource = tbl
        grdASTEXPT1.Text = "Comparative Retail Sales"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE IS NOT NULL", "SEQUENCE")
            Dim C As String = "G" & CStr(rowASTDSQLA.Item("SEQUENCE"))
            Dim D As String = rowASTDSQLA.Item("COLUMN_CAPTION")
            Set_DX_Column(grdASTEXPT1, C, D, 80)
        Next

        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim C As String = rowASTDSQLS.Item("COLUMN_NAME")
            Dim D As String = rowASTDSQLS.Item("COLUMN_CAPTION")
            Set_DX_Column(grdASTEXPT1, C, D, 100, "#,##0", "Sum")
        Next
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
        End If
    End Sub
  
    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        ASCDATA1.ExecuteSQL("Alter Table " & TT & " Add ITEM_CODE VARCHAR2(25)")

        Dim G As Int16 = COLUMN_NAMEs.Count
        If COLUMN_NAMEs(G - 1) = "ITEM_CODE" Then

            ASCMAIN1.sql = "Update " & TT & " Set ITEM_CODE = SUBSTR(G" & CStr(G) & "," & CStr(Len(COLUMN_CAPTIONs(G - 1)) + 2) & ")"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()
        Prepare_Custom_XLS()
    End Sub

    Sub SubTotal(THIS_LEVEL As Integer, _
                 LAST_LEVEL As Integer, _
                 G_Colors() As System.Drawing.Color, _
                 ByRef XR As Integer, _
                 ByRef XC As Integer, _
                 ByRef JMIN As Integer, _
                 ByRef JMAX As Integer, _
                 XWS As Microsoft.Office.Interop.Excel.Worksheet, _
                 keysG() As String, _
                 CURRENT_VALUE() As String, _
                 STL As Dictionary(Of Integer, List(Of Integer)), _
                 recaps() As Dictionary(Of String, List(Of Integer)), _
                 PIVOT_VALUE_COL_FORMULAS() As String)

        'Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim r As String

        If LAST_LEVEL <> -1 Then
            For ST As Integer = keysG.Length - 1 To THIS_LEVEL Step -1

                Dim RC As Dictionary(Of String, List(Of Integer)) = recaps(ST)
                Dim RC_VALUE As String = CURRENT_VALUE(ST)
                If ST <> 0 Then
                    If Not RC.ContainsKey(RC_VALUE) Then
                        RC.Add(RC_VALUE, New List(Of Integer))
                    End If
                End If

                XR += 1
                If ST = 0 Then
                    XWS.Cells(XR, XC + 1).VALUE = "Grand Totals"
                Else
                    For i As Integer = ST To keysG.Length - 1
                        XWS.Cells(XR, XC + ST).VALUE = CURRENT_VALUE(ST) & " Totals"
                        ' why are we looping here - isnt this the totals line?
                    Next
                End If
                Dim STFX As String = ""
                For Each II As Integer In STL(ST + 1)
                    STFX &= "," & ":" & CStr(II)
                Next

                If STFX <> "" Then
                    Dim STF As String = "=@SUM(" & Mid(STFX, 2) & ")"
                    Dim pvi As Integer = 0
                    Dim pvc As Integer = 0
                    Dim pvx As Integer = 0
                    For J As Integer = JMIN To JMAX
                        pvx += 1
                        pvi = 1 + (pvx - 1) \ PIVOT_VALUE_COL_FORMULAS.Length
                        pvc = 1 + ((pvx - 1) Mod PIVOT_VALUE_COL_FORMULAS.Length)
                        Dim XCC As String = Split(Me.Excel_Cell(0, J), ":")(0)
                        If PIVOT_VALUE_COL_FORMULAS(pvc - 1) <> "" Then
                            Dim CFx As String = PIVOT_VALUE_COL_FORMULAS(pvc - 1)
                            For I As Integer = 1 To PIVOT_VALUE_COL_FORMULAS.Length
                                'CFx = Replace(CF, Chr(Asc("A") + I - 1) & "#", Chr(Asc("A") + I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length) & CStr(XR))
                                Dim xcx As String = Split(Excel_Cell(0, I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length), ":")(0)
                                CFx = Replace(CFx, Chr(Asc("A") + I - 1) & "#", xcx & CStr(XR))
                            Next

                            XWS.Cells(XR, J).VALUE = "=" & CFx
                        Else
                            XWS.Cells(XR, J).VALUE = Replace(STF, ":", XCC)
                        End If
                    Next
                    r = Me.Excel_Cell(XR, IIf(ST = 0, 1, ST)) & ":" & Me.Excel_Cell(XR, JMAX)
                    With XWS.Range(r)
                        '.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        '.Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        '.Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                        '.Interior.TintAndShade = -0.14996795556505
                        '.Interior.PatternTintAndShade = 0

                        .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                        .Interior.Color = System.Drawing.ColorTranslator.ToOle(G_Colors(ST))
                        '.Interior.TintAndShade = -0.13 * ST
                        .Interior.PatternTintAndShade = 0
                        '.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue)
                        .Font.Bold = True
                    End With


                    r = Me.Excel_Cell(XR, 1) & ":" & Me.Excel_Cell(XR, JMAX)
                    With XWS.Range(r)
                        .BorderAround()
                    End With

                End If

                STL(ST + 1) = New List(Of Integer)
                If ST <> 0 Then STL(ST).Add(XR)

                If ST <> 0 Then
                    RC(RC_VALUE).Add(XR)
                End If
            Next
        End If

    End Sub

    Sub Prepare_Custom_XLS_Heading(oSheet As SpreadsheetGear.IWorksheet)
        ' Worksheet Heading

        With oSheet.Cells(0, 0)
            ' .Value = Format(Now, "MM/dd/yyyy HH:mm")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .NumberFormat = "mm/dd/yy;@"
            .Value = Now
        End With
        With oSheet.Cells(0, 1)
            .Value = MENU_ITEM_OBJECT
        End With
        With oSheet.Cells(0, 2)
            .Value = ASCMAIN1.USER_ID
        End With
        With oSheet.Cells(1, 0)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Font.Name = "Times New Roman"
            ' .Name = "Verdana"
            .Value = MENU_ITEM_DESC
        End With
        With oSheet.Cells(2, 0)
            .Font.Color = SpreadsheetGear.Colors.Blue
            .Font.Size = 20
            .Value = SUBT
        End With
    End Sub

    Sub Prepare_Custom_XLS(Optional ByVal xls_where As String = "", Optional ByVal ASTSRPT1 As String = "ASTSRPT1")

        ' Declare SSG Objects

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        ' Parameters

        Dim Heading_Anchor_Row As Integer = 5

        ' Save Workbook as FILENAME

        Dim FILENAME_TEMPLATE As String = ""
        Dim FILENAME_SOURCE As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & FILENAME_TEMPLATE
        Dim XLS_FILENAME As String = ""

        ASCMAIN1.Progress("Now Creating Custom XLS Workbook")
        If FILENAME_TEMPLATE = "" Then
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            'oSheet = oWB.Worksheets.Add
            'oSheet.Name = "Data"
            XLS_FILENAME = ASCMAIN1.Folders("Work") & XNO & ".xlsx"
            oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Else
            Dim success As Boolean = False
            Dim XLS_NO As Integer = 0

            Do Until success
                Try
                    XLS_NO += 1
                    XLS_FILENAME = ASCMAIN1.Folders("Work") & "Comparative_Retail_Sales"
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                    FileCopy(FILENAME_SOURCE, XLS_FILENAME)
                    success = True

                Catch ex As Exception
                    ' Stop
                End Try
            Loop

            oWB = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)
            oSheet = oWB.Worksheets("Data")
        End If


        Dim XTD_colors() As SpreadsheetGear.Color = _
        {SpreadsheetGear.Colors.PaleTurquoise, _
         SpreadsheetGear.Colors.PaleGoldenrod, _
         SpreadsheetGear.Colors.PaleGreen, _
         SpreadsheetGear.Colors.Beige}

        Dim G_Colors(9) As SpreadsheetGear.Color
        G_Colors(1) = SpreadsheetGear.Colors.Purple
        G_Colors(2) = SpreadsheetGear.Colors.Green
        G_Colors(3) = SpreadsheetGear.Colors.DarkOrange
        G_Colors(4) = SpreadsheetGear.Colors.Blue
        G_Colors(5) = SpreadsheetGear.Colors.Olive
        G_Colors(6) = SpreadsheetGear.Colors.Brown
        G_Colors(7) = SpreadsheetGear.Colors.Gold
        G_Colors(8) = SpreadsheetGear.Colors.DarkMagenta
        G_Colors(9) = SpreadsheetGear.Colors.Red


        Dim COLs_List As New List(Of String)

        Dim YTD As String = ""
        Dim STD_S As String = ""
        Dim STD_F As String = ""
        Dim QTD_1 As String = ""
        Dim QTD_2 As String = ""
        Dim QTD_3 As String = ""
        Dim QTD_4 As String = ""
        For M As Integer = 1 To 12
            If Format(M, "00") <= Mid(RYP, 5, 2) Then
                YTD &= "+TY_M" & Format(M, "00")
                If M <= 6 Then STD_S &= "+TY_M" & Format(M, "00")
                If M >= 7 Then STD_F &= "+TY_M" & Format(M, "00")
                If M >= 1 And M <= 3 Then QTD_1 &= "+TY_M" & Format(M, "00")
                If M >= 4 And M <= 6 Then QTD_2 &= "+TY_M" & Format(M, "00")
                If M >= 7 And M <= 9 Then QTD_3 &= "+TY_M" & Format(M, "00")
                If M >= 10 And M <= 12 Then QTD_4 &= "+TY_M" & Format(M, "00")
            End If
        Next
        dst.Tables("ASTSRPT1").Columns.Add("TY_M00", GetType(System.Decimal), Mid(YTD, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LY_M00", GetType(System.Decimal), Mid(Replace(YTD, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TY_B00", GetType(System.Decimal), Mid(Replace(YTD, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_S", GetType(System.Decimal), Mid(STD_S, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_S", GetType(System.Decimal), Mid(Replace(STD_S, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_S", GetType(System.Decimal), Mid(Replace(STD_S, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_F", GetType(System.Decimal), Mid(STD_F, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_F", GetType(System.Decimal), Mid(Replace(STD_F, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_F", GetType(System.Decimal), Mid(Replace(STD_F, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_Q1", GetType(System.Decimal), Mid(QTD_1, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_Q1", GetType(System.Decimal), Mid(Replace(QTD_1, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_Q1", GetType(System.Decimal), Mid(Replace(QTD_1, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_Q2", GetType(System.Decimal), Mid(QTD_2, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_Q2", GetType(System.Decimal), Mid(Replace(QTD_2, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_Q2", GetType(System.Decimal), Mid(Replace(QTD_2, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_Q3", GetType(System.Decimal), Mid(QTD_3, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_Q3", GetType(System.Decimal), Mid(Replace(QTD_3, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_Q3", GetType(System.Decimal), Mid(Replace(QTD_3, "TY_M", "TY_B"), 2))

        dst.Tables("ASTSRPT1").Columns.Add("TYA_Q4", GetType(System.Decimal), Mid(QTD_4, 2))
        dst.Tables("ASTSRPT1").Columns.Add("LYA_Q4", GetType(System.Decimal), Mid(Replace(QTD_4, "TY_M", "LY_M"), 2))
        dst.Tables("ASTSRPT1").Columns.Add("TYB_Q4", GetType(System.Decimal), Mid(Replace(QTD_4, "TY_M", "TY_B"), 2))


        Dim Blocks As New List(Of String)

        Add_Blocks(Blocks, "Total", COLs_List, New String() {"", "LY_M00", "TY_B00", "TY_V_PL", "TY_M00", "TY_V_LY"})
        Add_Blocks(Blocks, "Spring", COLs_List, New String() {"", "LYA_S", "TYB_S", "TY_V_PL_S", "TYA_S", "TY_V_LY_S"})
        Add_Blocks(Blocks, "Fall", COLs_List, New String() {"", "LYA_F", "TYB_F", "TY_V_PL_F", "TYA_F", "TY_V_LY_F"})
        Add_Blocks(Blocks, "Qtr 1 JAS", COLs_List, New String() {"", "LYA_Q1", "TYB_Q1", "TY_V_PL_Q1", "TYA_Q1", "TY_V_LY_Q1"})
        Add_Blocks(Blocks, "Qtr 2 OND", COLs_List, New String() {"", "LYA_Q2", "TYB_Q2", "TY_V_PL_Q2", "TYA_Q2", "TY_V_LY_Q2"})
        Add_Blocks(Blocks, "Qtr 3 JFM", COLs_List, New String() {"", "LYA_Q3", "TYB_Q3", "TY_V_PL_Q3", "TYA_Q3", "TY_V_LY_Q3"})
        Add_Blocks(Blocks, "Qtr 4 AMJ", COLs_List, New String() {"", "LYA_Q4", "TYB_Q4", "TY_V_PL_Q4", "TYA_Q4", "TY_V_LY_Q4"})

        For M As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, M - 1)
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
            Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
            Add_Blocks(Blocks, LEGEND, COLs_List, New String() {"", _
                                                                "LY_M" & Format(M, "00"), _
                                                                "TY_B" & Format(M, "00"), _
                                                                "TY_M" & Format(M, "00"), _
                                                                "TYB" & Format(M, "00"), _
                                                                "TYL" & Format(M, "00")})
        Next

        'For M As Integer = 1 To 12
        '    COLs_List.Add("")
        '    COLs_List.Add("LY_M" & Format(M, "00"))
        '    COLs_List.Add("TY_B" & Format(M, "00"))
        '    COLs_List.Add("TY_M" & Format(M, "00"))
        '    COLs_List.Add("TYB" & Format(M, "00"))
        '    COLs_List.Add("TYL" & Format(M, "00"))
        'Next

        Dim COLs() As String = COLs_List.ToArray


        ' Prepare to Traverse Dataset

  
        Dim FS As New Dictionary(Of String, String)
        Dim XLC As New Dictionary(Of String, String)



        Dim C As Integer = 0
        Dim R As Integer = 0
        Dim GMAX As Integer = COLUMN_NAMEs.Count



        Dim DC As New List(Of DataColumn)
        For i As Integer = 1 To GMAX
            DC.Add(dst.Tables("ASTSRPT1").Columns("G" & CStr(i)))
        Next
        DC.Add(dst.Tables("ASTSRPT1").Columns("INV_TYPE"))
        dst.Tables("ASTSRPT1").PrimaryKey = DC.ToArray



        C = GMAX + 1 + 1
        For Each SCN As String In COLs
            C += 1

            If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                If dst.Tables("ASTSRPT1").Columns(SCN).Expression <> "" Then
                    Dim FORMULA As String = "=" & Replace(dst.Tables("ASTSRPT1").Columns(SCN).Expression, "IIF", "IF")
                    FS.Add(SCN, FORMULA)
                End If
            End If

            Dim CP As Integer = (C - 1) \ 26
            Dim XL As String = Chr(64 + C - CP * 26)
            If CP > 0 Then
                XL = Chr(64 + CP) & XL
            End If
            If SCN <> "" Then XLC.Add(SCN, XL & "#")
        Next


        Dim XL1 As Integer = 0
        Dim XL2 As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""

        Dim G() As String = Nothing
        Dim GK() As String = Nothing
        Dim B As Integer = 0
        Dim ST() As String = Nothing

        Dim GS As String = ""
        For I As Integer = 1 To GMAX
            GS &= "," & "G" & CStr(I)
        Next

        Dim sqlw As String = ""

        If xls_where <> "" Then
            If sqlw = "" Then
                sqlw = xls_where
            Else
                sqlw &= " and " & xls_where
            End If
        End If




        ' Summary Sheet
        Dim G_Sheets As Integer = 2

        Dim order_by As String = "" '  Join(P, ",")
        Dim Ps As New List(Of String)
        For i As Integer = 1 To G_Sheets
            Ps.Add("G" & CStr(i))
            order_by &= "," & "G" & CStr(i)
        Next
        Dim P() As String = Ps.ToArray


        oSheet = oWB.Worksheets(0)
        oSheet.Name = "Summary"
        Prepare_Custom_XLS_Heading(oSheet)

        R = Heading_Anchor_Row + 1 : Prepare_Custom_XLS_Group_Headings(R, G_Sheets, oSheet)
        R += 1

        Dim tbl As DataTable = ASCDATA1.SelectDistinct("ASTSRPT1", P)
        ' Dim order_by As String = Join(P, ",")

        For Each row As DataRow In tbl.Select("", Mid(order_by, 2))

            Dim Sheet_Name As String = ""
            Dim sqlwx As String = ""
            For J As Integer = 1 To G_Sheets
                GROUP_KEY = row.Item("G" & CStr(J)) & ""
                rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)
                Sheet_Name &= " " & rowASTGROUP.Item("GROUP_CODE")
                sqlwx &= " and G" & CStr(J) & " = '" & row.Item("G" & CStr(J)) & "" & "'"
            Next
            Sheet_Name = Mid(Sheet_Name, 2)

            For I As Integer = 1 To G_Sheets

                If G Is Nothing OrElse GK(I) <> row.Item("G" & CStr(I)) & "" Then
                    B = I

                    If G Is Nothing Then
                        ReDim G(GMAX)
                        ReDim GK(GMAX)
                        ReDim ST(GMAX)
                    End If


                    For J As Integer = B To G_Sheets

                        GROUP_KEY = row.Item("G" & CStr(J)) & ""
                        rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)

                        GK(J) = GROUP_KEY
                        G(J) = rowASTGROUP.Item("GROUP_CODE")
                        GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                        R += 1 ' HEADING

                        oSheet.Cells(R - 1, G_Sheets).Value = GROUP_DESC

                        oSheet.Hyperlinks.Add(oSheet.Cells(R - 1, G_Sheets), _
                                                  "", _
                                                  "'" & Sheet_Name & "'!A4", _
                                                 "Click Here to Navigate to " & Sheet_Name, _
                                                  GROUP_DESC)

                        oSheet.Cells(R - 1, 0).EntireRow.OutlineLevel = GMAX
                        oSheet.Cells(R - 1, G_Sheets).IndentLevel = J - 1
                        If J <> GMAX Then
                            oSheet.Cells(R - 1, G_Sheets).Font.Color = G_Colors(J)
                        End If

                        For C = 1 To J
                            oSheet.Cells(R - 1, C - 1).Value = G(C)
                            If C <> GMAX Then
                                oSheet.Cells(R - 1, C - 1).Font.Color = G_Colors(C)
                            End If
                        Next

                        If J <> GMAX Then
                            oSheet.Cells(R - 1, 0).EntireRow.Font.Color = G_Colors(J)
                        End If
                    Next
                End If
            Next




            sqlwx = Mid(sqlwx, 5) : If xls_where <> "" Then sqlwx &= " and " & xls_where
            Prepare_Custom_XLS_Data_Sheet(oWB, sqlwx, Sheet_Name, _
                Blocks, row, Heading_Anchor_Row, GMAX, GS, COLs, FS, XLC, XTD_colors, G_Colors)

        Next

         


        ' Save Document and Show

        oWB.Save()
        Show_Document(XLS_FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Prepare_Custom_XLS_Group_Headings( _
        R As Integer, _
        GX As Integer, _
        oSheet As SpreadsheetGear.IWorksheet, _
        Optional show_Data_Headings As Boolean = False, _
        Optional COLs() As String = Nothing, _
        Optional Blocks As List(Of String) = Nothing, _
        Optional XTD_colors() As SpreadsheetGear.Color = Nothing)

        Dim C As Integer = 0

        For C = 1 To GX
            With oSheet.Cells(R - 1, C - 1)
                .Value = COLUMN_CAPTION_by_Lvl(C)
                .EntireColumn.ColumnWidth = 10
                .EntireColumn.NumberFormat = "@"
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            End With
        Next

        With oSheet.Cells(R - 1, C - 1)
            .Value = "Description"
            .ColumnWidth = 30
        End With

        If show_Data_Headings Then

            Dim GMAX As Integer = GX

            C += 1
            'oSheet.Cells(R - 2, 0, R - 1, GMAX + 1 + COLs.Length).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
            oSheet.Cells(R - 2, 0, R - 1, GMAX + 1 + COLs.Length).Interior.Color = SpreadsheetGear.Colors.DarkGray
            oSheet.Cells(R - 2, 0, R - 1, GMAX + 1 + COLs.Length).Font.Color = SpreadsheetGear.Colors.White
            ' oSheet.Cells(R - 2, C - 1).Interior.Pattern = SpreadsheetGear.Pattern.Solid
            oSheet.Cells(R - 3, GMAX + 1, R - 1, GMAX + 1 + COLs.Length).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            oSheet.Cells(R - 3, GMAX + 1, R - 1, GMAX + 1 + COLs.Length).VerticalAlignment = SpreadsheetGear.VAlign.Center


            Dim Col_Width As Integer = 10
            '     Dim Start_Col As New Dictionary(Of String, Integer)

            Dim FMT As String = "#,##0"



            For Each Block As String In Blocks

                C += 1 : With oSheet.Cells(R - 1, C - 1) : .ColumnWidth = 1 : End With
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "LY" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = Col_Width : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()

                oSheet.Cells(R - 3, C - 1).Value = Block
                Dim clr As Integer = 0
                If Block = "Totals" Then
                    clr = 1
                ElseIf Block = "Spring" Or Block = "Fall" Then
                    clr = 2
                ElseIf Block.StartsWith("Qtr ") Then
                    clr = 3
                End If
                oSheet.Cells(R - 3, C - 1).Interior.Color = XTD_colors(clr)

                oSheet.Cells(R - 3, C - 1, R - 3, C + 4 - 1).Merge()

                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "Plan" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = Col_Width : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "TY vs" & vbCrLf & "Plan" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = Col_Width * 0.6 : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "TY" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = Col_Width : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "TY vs" & vbCrLf & "LY" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = Col_Width * 0.6 : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()

            Next
        End If

    End Sub
    Sub Prepare_Custom_XLS_Data_Sheet( _
               oWB As SpreadsheetGear.IWorkbook, _
               sqlw As String, _
               Sheet_Name As String,
               Blocks As List(Of String), _
               rowG As DataRow, _
               Heading_Anchor_Row As Integer, _
               GMAX As Integer, _
               Gs As String, _
               COLs() As String, _
               FS As Dictionary(Of String, String), _
               XLC As Dictionary(Of String, String), _
               XTD_colors() As SpreadsheetGear.Color, _
               G_Colors() As SpreadsheetGear.Color)


        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        Dim C As Integer = 0

        ' Data Sheets

        Dim XL1 As Integer = 0
        Dim XL2 As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""

        Dim G() As String = Nothing
        Dim GK() As String = Nothing
        Dim B As Integer = 0
        Dim ST() As String = Nothing

        Dim R As Integer = Heading_Anchor_Row + 1

        oSheet = oWB.Worksheets.Add
        Prepare_Custom_XLS_Heading(oSheet)
        oSheet.Name = Sheet_Name

        R = Heading_Anchor_Row
        R += 1 : Prepare_Custom_XLS_Group_Headings(R, GMAX, oSheet, True, COLs, Blocks, XTD_colors)


        oSheet.Hyperlinks.Add(oSheet.Cells(3, 0), _
                                  "", _
                                  "'" & "Summary" & "'!A1", _
                                 "Click Here to Navigate Back to Summary Sheet", _
                                  "Summary")
        'If ASCMAIN1.Running_in_VS And InStr(sqlw, "NORDSTROM") <> 0 Then Stop

        Dim order_by As String = "" '  Join(P, ",")
        Dim Ps As New List(Of String)
        For i As Integer = 1 To GMAX
            Ps.Add("G" & CStr(i))
            order_by &= "," & "G" & CStr(i)
        Next
        Dim P() As String = Ps.ToArray

        Dim rowC As Integer = 0

        Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1").Select(sqlw), P)
        For Each row As DataRow In tbl.Select("", Mid(order_by, 2))
            '  For Each row As DataRow In dst.Tables("ASTSRPT1").Select(sqlw, Mid(Gs, 2))

            rowC += 1

            Dim k() As String
            ReDim k(GMAX)

            For I As Integer = 1 To GMAX

                k(I - 1) = row.Item("G" & CStr(I))

                If G Is Nothing OrElse GK(I) <> row.Item("G" & CStr(I)) & "" Then
                    B = I

                    If G Is Nothing Then
                        ReDim G(GMAX)
                        ReDim GK(GMAX)
                        ReDim ST(GMAX)
                    Else
                        If B < GMAX Then
                            Prepare_Custom_XLS_SubTotals(B, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, oSheet)
                            XL1 = 0
                            XL2 = 0
                        End If
                    End If

                    For J As Integer = B To GMAX
                        GROUP_KEY = row.Item("G" & CStr(J)) & ""
                        rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)

                        GK(J) = GROUP_KEY
                        G(J) = rowASTGROUP.Item("GROUP_CODE")
                        GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                        R += 1 ' HEADING


                        oSheet.Cells(R - 1, GMAX).Value = GROUP_DESC
                        oSheet.Cells(R - 1, 0).EntireRow.OutlineLevel = GMAX
                        oSheet.Cells(R - 1, GMAX).IndentLevel = J - 1
                        If J <> GMAX Then
                            oSheet.Cells(R - 1, GMAX).Font.Color = G_Colors(J)
                        End If

                        For C = 1 To J
                            oSheet.Cells(R - 1, C - 1).Value = G(C)
                            If C <> GMAX Then
                                oSheet.Cells(R - 1, C - 1).Font.Color = G_Colors(C)
                            End If
                        Next

                        If J <> GMAX Then
                            oSheet.Cells(R - 1, 0).EntireRow.Font.Color = G_Colors(J)
                        End If
                    Next
                End If
            Next

            Prepare_XLS_Prepare_row(row)

            Dim R_Start As Integer = R

            For Each DT As String In New String() {"1", "2", "3", "4"}
                C = GMAX + 1 + 1
                If DT <> "1" Then R += 1
                Dim DT_DESC As String = "Net"
                If DT = "1" Then DT_DESC = "Gross Basic"
                If DT = "2" Then DT_DESC = "Gift Sets"
                If DT = "3" Then DT_DESC = "Returns"

                oSheet.Cells(R - 1, C - 1).Value = DT_DESC

                k(GMAX) = DT

                Dim rowDT As DataRow = dst.Tables("ASTSRPT1").Rows.Find(k)
                If rowDT IsNot Nothing Then

                    For Each SCN As String In COLs
                        C += 1
                        If FS.ContainsKey(SCN) Then
                            Dim FORMULA As String = FS(SCN)

                            For ISCN As Integer = COLs.Length - 1 To 0 Step -1
                                Dim SCN2 As String = COLs(ISCN)
                                If SCN2 <> "" Then
                                    If InStr(FORMULA, SCN2) <> 0 Then
                                        FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                                    End If
                                End If
                            Next

                            FORMULA = Replace(FORMULA, "#", CStr(R))
                            oSheet.Cells(R - 1, C - 1).Formula = FORMULA

                        Else
                            If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                                oSheet.Cells(R - 1, C - 1).Value = rowDT.Item(SCN)
                            End If
                        End If
                    Next
                End If



                Try
                    oSheet.Cells(R - 1, 0).EntireRow.OutlineLevel = GMAX
                    If XL1 = 0 Then XL1 = R
                    XL2 = R

                Catch ex As Exception
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
                End Try

            Next

            Prepare_Custom_XLS_Border(oSheet, R_Start - 1, GMAX + 1, R - 1, GMAX + 1 + COLs.Length)
            If rowC Mod 2 = 0 Then
                oSheet.Cells(R_Start - 1, GMAX + 1, R - 1, GMAX + 1 + COLs.Length).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
            End If
        Next



        If ST Is Nothing Then Exit Sub

        Prepare_Custom_XLS_SubTotals(0, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, oSheet)


        Dim Start_Col As Integer = GMAX + 1 + 1

        For Each Block As String In Blocks

            Start_Col += 1
            Prepare_Custom_XLS_Border(oSheet, Heading_Anchor_Row - 1, Start_Col, R - 1, Start_Col + 5 - 1)
            Prepare_Custom_XLS_Border(oSheet, Heading_Anchor_Row - 1, Start_Col, R - 1, Start_Col + 1 - 1)
            Prepare_Custom_XLS_Border(oSheet, Heading_Anchor_Row - 1, Start_Col + 3, R - 1, Start_Col + 5 - 1)
            Prepare_Custom_XLS_Border(oSheet, Heading_Anchor_Row - 2, Start_Col, Heading_Anchor_Row - 2, Start_Col + 5 - 1)
            Prepare_Custom_XLS_Border(oSheet, Heading_Anchor_Row - 1, Start_Col, Heading_Anchor_Row - 0, Start_Col + 5 - 1)

            Start_Col += 5
        Next

        oSheet.WindowInfo.DisplayGridlines = False

    End Sub
    Sub Prepare_Custom_XLS_Border(oSheet As SpreadsheetGear.IWorksheet, R1 As Int64, C1 As Int64, R2 As Int64, C2 As Int64)
        With oSheet.Range(R1, C1, R2, C2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With
    End Sub

    Sub Prepare_Custom_XLS_SubTotals( _
    ByVal B As Integer, _
    ByRef R As Integer, _
    ByVal GMAX As Integer, _
    ByVal XL1 As Integer, _
    ByVal XL2 As Integer, _
    ByVal ST() As String, _
    ByVal G() As String, _
    ByVal GK() As String, _
    ByVal COLs() As String, _
    ByVal FS As Dictionary(Of String, String), _
    ByVal XLC As Dictionary(Of String, String), _
    ByVal G_Colors() As SpreadsheetGear.Color, _
    ByVal ws As SpreadsheetGear.IWorksheet)

        Dim C As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""

        For Slvl As Integer = GMAX - 1 To B Step -1

            R += 1 ' SUB-TOTAL
            ws.Cells(R - 1, 0).EntireRow.Font.Color = G_Colors(Slvl)
            For J As Integer = Slvl To 1 Step -1
                ws.Cells(R - 1, J - 1).Value = G(J)
                ws.Cells(R - 1, J - 1).Font.Color = G_Colors(J)
            Next

            ST(Slvl) &= ",X" & CStr(R)

            If Slvl = 0 Then
                GROUP_DESC = "Totals"
            Else
                GROUP_KEY = GK(Slvl)
                rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)
                GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                ws.Cells(R - 1, GMAX).IndentLevel = Slvl - 1
            End If
            ws.Cells(R - 1, GMAX).Value = GROUP_DESC
            ws.Cells(R - 1, GMAX).Font.Color = G_Colors(Slvl)

            C = GMAX + 1 + 1
            For Each SCN As String In COLs
                C += 1
                Dim CP As Integer = (C - 1) \ 26
                Dim XL As String = Chr(64 + C - CP * 26)
                If CP > 0 Then
                    XL = Chr(64 + CP) & XL
                End If

                If SCN <> "" Then

                    If FS.ContainsKey(SCN) Then
                        Dim FORMULA As String = FS(SCN)
                        For ISCN As Integer = COLs.Length - 1 To 0 Step -1
                            Dim SCN2 As String = COLs(ISCN)
                            If SCN2 <> "" Then
                                If InStr(FORMULA, SCN2) <> 0 Then
                                    FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                                End If
                            End If
                        Next
                        'For Each SCN2 As String In COLs
                        '    If InStr(FORMULA, SCN2) <> 0 Then
                        '        FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                        '    End If
                        'Next
                        FORMULA = Replace(FORMULA, "#", CStr(R))
                        If GK(Slvl) = aRC Then
                        Else
                            ws.Cells(R - 1, C - 1).Formula = FORMULA
                        End If

                    Else
                        If GK(Slvl) = aRC Then
                        Else
                            If Slvl = GMAX - 1 Then
                                ws.Cells(R - 1, C - 1).Formula = "=SUM(" & XL & XL1 & ":" & XL & XL2 & ")"
                            Else
                                ws.Cells(R - 1, C - 1).Formula = "=SUM(" & Replace(Mid(ST(Slvl + 1), 2), "X", XL) & ")"
                            End If
                        End If
                    End If

                    If Slvl > 0 Then ws.Cells(R - 1, 0).EntireRow.OutlineLevel = Slvl
                    ws.Cells(R - 1, C - 1).Font.Color = G_Colors(Slvl)
                End If
            Next
            ST(Slvl + 1) = ""

            Dim CC As SpreadsheetGear.Color = SpreadsheetGear.Colors.PaleGoldenrod
            If Slvl = 0 Then CC = SpreadsheetGear.Colors.LightGray
            CC = SpreadsheetGear.Colors.WhiteSmoke
            For C = 1 To GMAX + 1 + 1 + COLs.Length
                ws.Cells(R - 1, C - 1).Interior.Color = CC
                ws.Cells(R - 1, C - 1).Interior.Pattern = SpreadsheetGear.Pattern.Solid
            Next
            R += 1
            ws.Cells(R - 1, 0).EntireRow.RowHeight = ws.Cells(R - 1, 0).EntireRow.Height * 0.25
        Next

    End Sub

    Sub Add_Blocks(Blocks As List(Of String), BLOCK_NAME As String, COLs_List As List(Of String), COLUMN_NAMEs() As String)
        Blocks.Add(BLOCK_NAME)
        For i As Integer = 0 To 5
            COLs_List.Add(COLUMN_NAMEs(i))
        Next
    End Sub
End Class