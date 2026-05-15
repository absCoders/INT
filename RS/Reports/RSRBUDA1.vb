Public Class RSRBUDA1

#Region "Declarations"
    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String

    Dim SEASON_YEAR_LY As String

    Dim RSTBUDA1 As String
#End Region

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
        'does this need to be removed?
        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            ASCMAIN1.Folders("SharedRoot") = "C:\SHARE\INT\"
        End If
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

        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")
        SEASON_YEAR = Mid(RYM, 1, 4)
        Dim MM As String = Mid(RYM, 5, 2)
        SEASON_YEAR_LY = Format(Val(Mid(RYM, 1, 4) - 1 * LYX), "0000")

        If Mid(RYM, 5, 2) >= "07" Then
            SEASON_TYPE = "F"
        Else
            SEASON_TYPE = "S"
        End If
        SEASON_CODE = SEASON_TYPE & SEASON_YEAR

        Dim M01 As String = IIf(SEASON_TYPE = "S", "01", "07")
        Dim M02 As String = IIf(SEASON_TYPE = "S", "02", "08")
        Dim M03 As String = IIf(SEASON_TYPE = "S", "03", "09")
        Dim M04 As String = IIf(SEASON_TYPE = "S", "04", "10")
        Dim M05 As String = IIf(SEASON_TYPE = "S", "05", "11")
        Dim M06 As String = IIf(SEASON_TYPE = "S", "06", "12")

        Dim sqlAuths As String = IIf(chkALLDOORS.Checked, "", vbCrLf _
            & "   and SATAUTH1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SATAUTH1.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and SATAUTH1.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_CLOSED IS NULL" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL")

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, BRAND_CODE, HC_CODE, COLLECTION_GENDER" & vbCrLf _
            & ", Sum (LY_M01) LY_M01, Sum (TY_M01) TY_M01, Sum (LY_M02) LY_M02, Sum (TY_M02) TY_M02, Sum (LY_M03) LY_M03, Sum (TY_M03) TY_M03, Sum (LY_M04) LY_M04, Sum (TY_M04) TY_M04, Sum (LY_M05) LY_M05, Sum (TY_M05) TY_M05, Sum (LY_M06) LY_M06, Sum (TY_M06) TY_M06" & vbCrLf _
            & ", Sum (LY_STL) LY_STL, Sum (TY_STL) TY_STL, Sum (LY_STD) LY_STD, Sum (TY_STD) TY_STD" & vbCrLf _
            & ", Sum (BUD_M01) BUD_M01, Sum (BUD_M02) BUD_M02, Sum (BUD_M03) BUD_M03, Sum (BUD_M04) BUD_M04, Sum (BUD_M05) BUD_M05, Sum (BUD_M06) BUD_M06" & vbCrLf _
            & ", Sum (BUD_STL) BUD_STL, Sum (BUD_STD) BUD_STD" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select X.* from " & IIf(chkALLDOORS.Checked, "", "SATAUTH1, ") & "ICTBRAN1, ICTCOLL0, (" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M01 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M01" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M01 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M01" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M02 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M02" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M02 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M02" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M03 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M03" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M03 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M03" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M04 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M04" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M04 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M04" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M05 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M05" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M05 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M05" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M06 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) LY_M06" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR & M06 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) TY_M06" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR_LY & M06 & "' then NVL(RSTRETL1.AMT_SOLD,0) else 0 End) LY_STL" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR & M01 & "' and '" & SEASON_YEAR & M06 & "' then NVL(RSTRETL1.AMT_SOLD,0) else 0 End) TY_STL" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR_LY & MM & "' then NVL(RSTRETL1.AMT_SOLD,0) else 0 End) LY_STD" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR & M01 & "' and '" & SEASON_YEAR & MM & "' then NVL(RSTRETL1.AMT_SOLD,0) else 0 End) TY_STD" & vbCrLf _
            & ", 0 BUD_M01, 0 BUD_M02, 0 BUD_M03, 0 BUD_M04, 0 BUD_M05, 0 BUD_M06" & vbCrLf _
            & ", 0 BUD_STL, 0 BUD_STD" & vbCrLf _
            & " from RSTRETL1, ICTITEM1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR & MM & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTRETL1") _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " union " & vbCrLf _
            & "Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", 0 LY_M01, 0 TY_M01, 0 LY_M02, 0 TY_M02, 0 LY_M03, 0 TY_M03, 0 LY_M04, 0 TY_M04, 0 LY_M05, 0 TY_M05, 0 LY_M06, 0 TY_M06" & vbCrLf _
            & ", 0 LY_STL, 0 TY_STL, 0 LY_STD, 0 TY_STD" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M01 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M01" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M02 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M02" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M03 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M03" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M04 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M04" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M05 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M05" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M06 & "',NVL(RSTBUDR1.BUDGET,0),0)) BUD_M06" & vbCrLf _
            & ", Sum (Case when RSTBUDR1.OPS_YYYYPP between '" & SEASON_YEAR & M01 & "' and '" & SEASON_YEAR & M06 & "' then NVL(RSTBUDR1.BUDGET,0) else 0 End) BUD_STL" & vbCrLf _
            & ", Sum (Case when RSTBUDR1.OPS_YYYYPP between '" & SEASON_YEAR & M01 & "' and '" & SEASON_YEAR & MM & "' then NVL(RSTBUDR1.BUDGET,0) else 0 End) BUD_STD" & vbCrLf _
            & " from RSTBUDR1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTBUDR1.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR & M06 & "'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTBUDR1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTBUDR1") _
            & " group by RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ") X " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_STATUS = 'A'" _
            & sqlAuths & vbCrLf _
            & ")  group by CUST_CODE, CUST_STORE_NO, BRAND_CODE, HC_CODE, COLLECTION_GENDER"

        RSTBUDA1 = ASCMAIN1.Temp_Table

        Dim rTotal As Int64
        Dim r0 As Integer

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
            & ", SOTSDSC1.SDS_NAME" & vbCrLf _
            & ", LY_M01, TY_M01, LY_M02, TY_M02, LY_M03, TY_M03, LY_M04, TY_M04, LY_M05, TY_M05, LY_M06, TY_M06" & vbCrLf _
            & ", LY_STL, TY_STL, LY_STD, TY_STD" & vbCrLf _
            & ", BUD_M01, BUD_M02, BUD_M03, BUD_M04, BUD_M05, BUD_M06" & vbCrLf _
            & ", BUD_STL, BUD_STD" & vbCrLf _
            & " from " & RSTBUDA1 & " X" & vbCrLf _
            & ",ICTBRAN1,ARTCUST1,ARTCUST2,SOTSELL1,SOTSREG1,ICTCOLL0,SOTSDSC1,SOTSELL1 SOTSELL1_AC" & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE (+) = X.HC_CODE" & vbCrLf _
            & "   and SOTSDSC1.SDS_CODE (+) = ARTCUST2.SDS_CODE" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   and SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 4
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_NAME,HC_NAME,CUST_STORE_NO")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":AJ" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=MainData!$A$" & CStr(r0) & ":$AL$" & CStr(r0 + DataTable.Rows.Count))
        ' wb.Names.Add("Data", "=MainData!$A$" & CStr(r0) & ":$AI$" & CStr(r0 + DataTable.Rows.Count))

        ASCMAIN1.Progress("-", "Formulas")

        xlSourceRange = ws.Range("AK" & CStr(r0 + 1) & ":AK" & CStr(r0 + 1))
        xlDestRange = ws.Range("AK" & CStr(r0 + 1) & ":AK" & CStr(r0 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        ws.Cells(1, 1).Value = Now
        Dim monthNum As Integer = Convert.ToInt32(MM)
        'A1 on all sheets
        If monthNum >= 1 AndAlso monthNum <= 6 Then
            ' January to June
            SEASON_TYPE = "S"
            ws.Cells(2, 1).Value = "Spring " & SEASON_YEAR & " Retail Bonus Goal vs Actual"
        ElseIf monthNum >= 7 AndAlso monthNum <= 12 Then
            ' July to December
            SEASON_TYPE = "F"
            ws.Cells(2, 1).Value = "Fall " & SEASON_YEAR & " Retail Bonus Goal vs Actual"
        End If

        'In fall we unhide presidents incentive 
        Dim presidentsIncentiveSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("President's Incentive")
        If SEASON_TYPE = "F" Then
            ' Unhide the sheet for Fall
            presidentsIncentiveSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible
        ElseIf SEASON_TYPE = "S" Then
            ' Keep the sheet hidden for Spring
            presidentsIncentiveSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
        End If

        'bydoorbymonth
        Dim byDoorByMonthSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("By door by month vs Goal")
        Dim ttlGoalByDoorSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("TTL Goal by Door")
        Dim specialtyBrandsGoalsSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("Specality Brands Goals")
        ' Define month names for Spring and Fall
        Dim springMonths As String() = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"}
        Dim fallMonths As String() = {"Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        ' Select appropriate month names based on SEASON_TYPE
        Dim selectedMonths As String() = If(SEASON_TYPE = "S", springMonths, fallMonths)

        Dim colIndex As Integer = 2 ' Starting from column B (2)
        For Each sm In selectedMonths
            byDoorByMonthSheet.Cells(8, colIndex).Value = sm & " MTD Actual"
            colIndex += 1
            byDoorByMonthSheet.Cells(8, colIndex).Value = sm & " Goal"
            colIndex += 1
            byDoorByMonthSheet.Cells(8, colIndex).Value = sm & " Goal $ to Go"
            colIndex += 1
            byDoorByMonthSheet.Cells(8, colIndex).Value = sm & " Goal Achieved %"
            colIndex += 1
        Next
        ' Update column headers from B8 to G8
        For i As Integer = 0 To 5
            ttlGoalByDoorSheet.Cells(8, i + 2).Value = selectedMonths(i) & " Goal"
            specialtyBrandsGoalsSheet.Cells(9, i + 5).Value = selectedMonths(i) & " Goal"
        Next

        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Retail_Sales_Budget_vs_Actual"
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