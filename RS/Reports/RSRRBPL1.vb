Public Class RSRRBPL1

    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String
    Dim SEASON_TYPE_PS As String
    Dim SEASON_YEAR_PS As String

    Dim RSTRBPL1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)
        cbeLY.SelectedIndex = 0
    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        Create_Pivot()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Sub Create_Pivot()

        Dim SQLW As String = ""
        SQLW &= SQLA_filter("CUST_CODE", "RSTRETL1")
        'SQLW &= SQLA_filter("SREP_CODE", "SOTORDR1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")

        Dim LYX As Integer = 1
        LYX = Val(cbeLY.Value & "")

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)


        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        Dim LYW As String = ASCMAIN1.Week_Calc(RYW, -52 * LYX)
        If ASCMAIN1.CLIENT = "INT" Then
            LYW = Format(Val(Mid(RYW, 1, 4)) - 1 * LYX, "0000") & Mid(RYW, 5, 2)
        End If

        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")
        'REL_WEEK = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        'Dim LEGEND_WK As String = rowGLTPARM3.Item("LEGEND")
        'WEEK_LEGEND = "Week Ending " & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yy") & ",  " & LEGEND_WK

        SEASON_YEAR_PS = Mid(RYM, 1, 4)

        Dim WKS As Integer = Val(Mid(RYW, 5, 2)) + 4
        If WKS > 52 Then WKS = WKS - 52 * LYX

        If Mid(RYM, 5, 2) >= "07" Then
            SEASON_TYPE_PS = "F"
            SEASON_TYPE = "S"
            SEASON_YEAR = Format(Val(SEASON_YEAR_PS) + 1, "0000")
            WKS = WKS - 26
        Else
            SEASON_TYPE_PS = "S"
            SEASON_TYPE = "F"
            SEASON_YEAR = SEASON_YEAR_PS
        End If

        Dim RYW_SWK1 As String = ASCMAIN1.Week_Calc(RYW, -WKS + 1)
        Dim LYW_SWK1 As String = ASCMAIN1.Week_Calc(RYW, -WKS + 1 - 52 * LYX)

        If ASCMAIN1.CLIENT = "INT" Then
            If Mid(RYW_SWK1, 5, 2) = "50" Then
                Mid(RYW_SWK1, 5, 2) = "49"
            End If
            If Mid(LYW_SWK1, 5, 2) = "50" Then
                Mid(LYW_SWK1, 5, 2) = "49"
            End If
            If Mid(RYW_SWK1, 5, 2) = "24" Then
                Mid(RYW_SWK1, 5, 2) = "23"
            End If
            If Mid(LYW_SWK1, 5, 2) = "24" Then
                Mid(LYW_SWK1, 5, 2) = "23"
            End If
        End If

        SEASON_CODE = SEASON_TYPE & SEASON_YEAR

        Dim M01 As String = IIf(SEASON_TYPE = "S", "01", "07")
        Dim M02 As String = IIf(SEASON_TYPE = "S", "02", "08")
        Dim M03 As String = IIf(SEASON_TYPE = "S", "03", "09")
        Dim M04 As String = IIf(SEASON_TYPE = "S", "04", "10")
        Dim M05 As String = IIf(SEASON_TYPE = "S", "05", "11")
        Dim M06 As String = IIf(SEASON_TYPE = "S", "06", "12")

        Dim PS_M01 As String = IIf(SEASON_TYPE = "S", "07", "01")
        Dim PS_M06 As String = IIf(SEASON_TYPE = "S", "12", "06")

        Dim SEASON_CODE_LY = Mid(SEASON_CODE, 1, 1) & Format(Val(Mid(SEASON_CODE, 2, 4)) - 1 * LYX, "0000")
        Dim SEASON_CODE_2LY = Mid(SEASON_CODE, 1, 1) & Format(Val(Mid(SEASON_CODE, 2, 4)) - 1 - 1 * LYX, "0000")

        Dim O As Integer = 0
        Dim SN2LY1 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 - 1 * LYX, "0000") & M01, O)
        Dim SN2LY6 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 - 1 * LYX, "0000") & M06, O)
        Dim SNLY1 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M01, O)
        Dim SNLY2 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M02, O)
        Dim SNLY3 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M03, O)
        Dim SNLY4 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M04, O)
        Dim SNLY5 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M05, O)
        Dim SNLY6 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE, 2)) - 1 * LYX, "0000") & M06, O)

        Dim SEASON_CODE_PRIOR As String = SEASON_TYPE_PS & SEASON_YEAR_PS ' IIf(SEASON_TYPE = "S", "F" & Format(Mid(SEASON_CODE, 2, 4) - 1, "0000"), "S" & Mid(SEASON_CODE, 2, 4))
        Dim SEASON_CODE_PRIOR_LY = Mid(SEASON_CODE_PRIOR, 1, 1) & Format(Mid(SEASON_CODE_PRIOR, 2, 4) - 1 * LYX, "0000")
        Dim PSN_M1 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE_PRIOR, 2)), "0000") & PS_M01, O)
        Dim PSN_M6 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE_PRIOR, 2)), "0000") & PS_M06, O)
        Dim PSN_LY_M1 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE_PRIOR_LY, 2)), "0000") & PS_M01, O)
        Dim PSN_LY_M6 As String = ASCMAIN1.Period_Calc(Format(Val(Mid(SEASON_CODE_PRIOR_LY, 2)), "0000") & PS_M06, O)

        Dim sqlAuths As String = IIf(chkALLDOORS.Checked, "", vbCrLf _
            & "   and SATAUTH1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SATAUTH1.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and SATAUTH1.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and (SATAUTH1.OPS_YYYYPP_CLOSED IS NULL OR SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL")

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select X.* from " & IIf(chkALLDOORS.Checked, "", "SATAUTH1, ") & "ICTBRAN1, ICTCOLL0, (" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SN2LY1 & "' and '" & SN2LY6 & "' Then NVL(RSTRETL1.AMT_SOLD,0) / 1 ELSE 0 END) SSN_2LY" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SNLY1 & "' and '" & SNLY6 & "' Then NVL(RSTRETL1.AMT_SOLD,0) / 1 ELSE 0 END) SSN_LY" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY1 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M01" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY2 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M02" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY3 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M03" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY4 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M04" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY5 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M05" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SNLY6 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) SNLY_M06" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & LYW_SWK1 & "' and '" & LYW & "' Then NVL(RSTRETL1.AMT_SOLD,0) / 1 ELSE 0 END) PSN_STD_LY" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & RYW_SWK1 & "' and '" & RYW & "' Then NVL(RSTRETL1.AMT_SOLD,0) / 1 ELSE 0 END) PSN_STD" & vbCrLf _
            & " from RSTRETL1, ICTITEM1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP between '" & SN2LY1 & "' and '" & PSN_M6 & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & SQLW _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ") X " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_STATUS = 'A'" _
            & sqlAuths

        RSTRBPL1 = ASCMAIN1.Temp_Table


        ASCMAIN1.Progress("-", "Lookups")


        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Customer / HC Summary")

        ws = wb.Worksheets("ChainPlanner")

        ASCMAIN1.sql = "Select Distinct " & vbCrLf _
            & "  ICTCOLL0.HC_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & " from " & RSTRBPL1 & " X" & vbCrLf _
            & ",ICTCOLL0,ARTCUST1" & vbCrLf _
            & " where ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 2
        rTotal = DataTable.Select("", "HC_NAME, CUST_NAME").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "HC_NAME,CUST_NAME")
            r += 1
            ws.Range("B" & CStr(r0 + r) & ":C" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 1000 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next


        For i As Integer = 4 To 14 ' 22
            If i <> 10 Then
                Dim cc3 As Integer = i
                xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc3))
                xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc3), Excel_Cell(r0 + DataTable.Rows.Count, cc3))
                xlSourceRange.Copy(xlDestRange)
            End If
        Next

        xlSourceRange = ws.Range(Excel_Cell(r0 + 1, 1))
        xlDestRange = ws.Range(Excel_Cell(r0 + 1, 1), Excel_Cell(r0 + DataTable.Rows.Count, 1))
        xlSourceRange.Copy(xlDestRange)


        wb.Names.Add("Key_2", "=CHAINPLANNER!$A$" & CStr(r0) & ":$K$" & CStr(r0 + DataTable.Rows.Count))


        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("MainData")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.HC_CODE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME" & vbCrLf _
            & ", SOTSREG1.REGION_DESC" & vbCrLf _
            & ", SOTSELL1.SELL_NAME" & vbCrLf _
            & ", SOTSELL1_AC.SELL_NAME AC" & vbCrLf _
            & ", X.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL0.HC_NAME" & vbCrLf _
            & ", DECODE(X.COLLECTION_GENDER,'W',ARTCUST2.CUST_STORE_VOL_RANK_W,'M',ARTCUST2.CUST_STORE_VOL_RANK_M,'?') STORE_RANK" & vbCrLf _
            & ", DECODE(X.COLLECTION_GENDER,'W',ARTCUST2.CUST_STORE_CUST_RANK_W,'M',ARTCUST2.CUST_STORE_CUST_RANK_M,'?') STORE_RANK_CUST" & vbCrLf _
            & ", NULL RANK_IN_STORE" & vbCrLf _
            & ", SSN_2LY" & vbCrLf _
            & ", SSN_LY" & vbCrLf _
            & ", PSN_STD_LY" & vbCrLf _
            & ", PSN_STD" & vbCrLf _
            & ", SNLY_M01, SNLY_M02, SNLY_M03, SNLY_M04, SNLY_M05, SNLY_M06" & vbCrLf _
            & " from " & RSTRBPL1 & " X" & vbCrLf _
            & ",ICTBRAN1,ARTCUST1,ARTCUST2,SOTSELL1,SOTSREG1,ICTCOLL0,SOTSELL1 SOTSELL1_AC" & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE (+) = X.HC_CODE" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   and SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 15
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_NAME,HC_NAME,CUST_STORE_NO")
            r += 1
            ws.Range("B" & CStr(r0 + r) & ":Q" & CStr(r0 + r)).Value2 = row.ItemArray

            ws.Cells(r0 + r, 22).Value = row.Item("PSN_STD_LY")
            ws.Cells(r0 + r, 23).Value2 = row.Item("PSN_STD")

            ws.Cells(r0 + r, 29).Value2 = row.Item("SNLY_M01")
            ws.Cells(r0 + r, 31).Value2 = row.Item("SNLY_M02")
            ws.Cells(r0 + r, 33).Value2 = row.Item("SNLY_M03")
            ws.Cells(r0 + r, 35).Value2 = row.Item("SNLY_M04")
            ws.Cells(r0 + r, 37).Value2 = row.Item("SNLY_M05")
            ws.Cells(r0 + r, 39).Value2 = row.Item("SNLY_M06")

            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=MAINDATA!$A$" & CStr(r0) & ":$AA$" & CStr(r0 + DataTable.Rows.Count))

        wb.Names.Add("KEY", "=MAINDATA!$A$" & CStr(r0) & ":$A$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CURR_SN", "=MAINDATA!$W$" & CStr(r0) & ":$W$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CURR_SN_LY", "=MAINDATA!$V$" & CStr(r0) & ":$V$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("Fall_14", "=MAINDATA!$Q$" & CStr(r0) & ":$Q$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("Fall_13", "=MAINDATA!$P$" & CStr(r0) & ":$P$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("FC_2015", "=MAINDATA!$Z$" & CStr(r0) & ":$Z$" & CStr(r0 + DataTable.Rows.Count))


        ws.Cells(r0 - 2, 16).Value = SEASON_CODE_2LY
        ws.Cells(r0 - 2, 17).Value = SEASON_CODE_LY

        ws.Cells(r0 - 2, 22).Value = SEASON_CODE_PRIOR_LY & " STD"
        ws.Cells(r0 - 2, 23).Value = SEASON_CODE_PRIOR & " STD"
        ws.Cells(r0 - 2, 24).Value = SEASON_CODE_PRIOR & " Trend"
        ws.Cells(r0 - 2, 25).Value = SEASON_CODE

        ws.Cells(r0 - 2, 41).Value = SEASON_CODE_LY & " Total"
        ws.Cells(r0 - 2, 42).Value = SEASON_CODE & " Total"

        ASCMAIN1.Progress("-", "Formulas")

        xlSourceRange = ws.Range("A" & CStr(r0 + 1) & ":A" & CStr(r0 + 1))
        xlDestRange = ws.Range("A" & CStr(r0 + 1) & ":A" & CStr(r0 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        Dim cc2 As Integer

        For i As Integer = 24 To 27
            If i = 26 Then
            Else
                cc2 = i
                xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc2))
                xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc2), Excel_Cell(r0 + DataTable.Rows.Count, cc2))
                xlSourceRange.Copy(xlDestRange)
            End If
        Next

        For i As Integer = 18 To 21
            'If i = 21 Then ' 18 thru 21 when sp sets formulsa
            cc2 = i
            xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc2))
            xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc2), Excel_Cell(r0 + DataTable.Rows.Count, cc2))
            xlSourceRange.Copy(xlDestRange)
            'End If
        Next

        For i As Integer = 1 To 6
            Dim cc As Integer = 30 + (i - 1) * 2
            xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc))
            xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc), Excel_Cell(r0 + DataTable.Rows.Count, cc))
            xlSourceRange.Copy(xlDestRange)
        Next

        For i As Integer = 1 To 2
            Dim cc As Integer = 40 + i
            xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc))
            xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc), Excel_Cell(r0 + DataTable.Rows.Count, cc))
            xlSourceRange.Copy(xlDestRange)
        Next

        ws.Cells(1, 1).Value = Now


        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")
       
        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Retail_Sales_Planner"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                ' Stop
            End Try
        Loop

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

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

    End Sub

End Class