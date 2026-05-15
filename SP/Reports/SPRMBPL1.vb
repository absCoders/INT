Public Class SPRMBPL1

    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String
    Dim SEASON_TYPE_PS As String
    Dim SEASON_YEAR_PS As String

    Dim SPTMBPL1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)
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
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")

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

        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")

        SEASON_YEAR_PS = Mid(RYM, 1, 4)
        ' IF YOU PICK A WEEK IN S, THEN THE PLANNER THINKS YOU WANT TO PLAN F
        ' IF YOU PICK A WEEK IN F, THEN THE PLANNER THINKS YOU WANT TO PLAN S
        ' NOTE THAT IF YOU PICK A WEEK IN JAN, THAT THE PLANNER THINKS THAT YOU WANT TO PLAN FALL

        Dim WKS As Integer = Val(Mid(RYW, 5, 2)) + 4
        If WKS > 52 Then WKS = WKS - 52

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
        SEASON_CODE = SEASON_YEAR & SEASON_TYPE

        Dim SEASON_YEAR_LY As String = Format(Val(SEASON_YEAR) - 1, "0000")

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
            & "   and (SATAUTH1.OPS_YYYYPP_CLOSED IS NULL OR SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL")

        Dim RSTBUDR1 As String = TAC.RSCMAIN1.RSTBUDR1_AS_YP

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select X.* from " & IIf(chkALLDOORS.Checked, "", "SATAUTH1, ") & "ICTBRAN1, ICTCOLL0, (" & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, BRAND_CODE, HC_CODE, COLLECTION_GENDER" & vbCrLf _
            & ", Sum (RS_SSN) RS_SSN, Sum (RS_M01) RS_M01, Sum (RS_M02) RS_M02, Sum (RS_M03) RS_M03, Sum (RS_M04) RS_M04, Sum (RS_M05) RS_M05, Sum (RS_M06) RS_M06" & vbCrLf _
            & ", Sum (MX_SSN) MX_SSN, Sum (MX_M01) MX_M01, Sum (MX_M02) MX_M02, Sum (MX_M03) MX_M03, Sum (MX_M04) MX_M04, Sum (MX_M05) MX_M05, Sum (MX_M06) MX_M06" & vbCrLf _
            & ", Sum (RB_SSN) RB_SSN, Sum (RB_M01) RB_M01, Sum (RB_M02) RB_M02, Sum (RB_M03) RB_M03, Sum (RB_M04) RB_M04, Sum (RB_M05) RB_M05, Sum (RB_M06) RB_M06" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", Sum (NVL(RSTRETL1.AMT_SOLD,0)) RS_SSN" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M01 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M01" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M02 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M02" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M03 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M03" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M04 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M04" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M05 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M05" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYPP,'" & SEASON_YEAR_LY & M06 & "',NVL(RSTRETL1.AMT_SOLD,0),0)) RS_M06" & vbCrLf _
            & ", 0 MX_SSN, 0 MX_M01, 0 MX_M02, 0 MX_M03, 0 MX_M04, 0 MX_M05, 0 MX_M06" & vbCrLf _
            & ", 0 RB_SSN, 0 RB_M01, 0 RB_M02, 0 RB_M03, 0 RB_M04, 0 RB_M05, 0 RB_M06" & vbCrLf _
            & " from RSTRETL1, ICTITEM1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR_LY & M06 & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTRETL1") _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", 0 RS_SSN, 0 RS_M01, 0 RS_M02, 0 RS_M03, 0 RS_M04, 0 RS_M05, 0 RS_M06" & vbCrLf _
            & ", Sum (NVL(SPTCWRX2.BILL_AMT,0)) MX_SSN" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M01 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M01" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M02 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M02" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M03 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M03" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M04 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M04" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M05 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M05" & vbCrLf _
            & ", Sum (Decode(SPTCWRX2.OPS_YYYYPP,'" & SEASON_YEAR_LY & M06 & "',NVL(SPTCWRX2.BILL_AMT,0),0)) MX_M06" & vbCrLf _
            & ", 0 RB_SSN, 0 RB_M01, 0 RB_M02, 0 RB_M03, 0 RB_M04, 0 RB_M05, 0 RB_M06" & vbCrLf _
            & " from SPTCWRX2, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where SPTCWRX2.OPS_YYYYPP between '" & SEASON_YEAR_LY & M01 & "' and '" & SEASON_YEAR_LY & M06 & "'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = SPTCWRX2.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = SPTCWRX2.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "SPTCWRX2") _
            & " group by SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " union " & vbCrLf _
            & "Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", 0 RS_SSN, 0 RS_M01, 0 RS_M02, 0 RS_M03, 0 RS_M04, 0 RS_M05, 0 RS_M06" & vbCrLf _
            & ", 0 MX_SSN, 0 MX_M01, 0 MX_M02, 0 MX_M03, 0 MX_M04, 0 MX_M05, 0 MX_M06" & vbCrLf _
            & ", 0 RB_SSN" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M01 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M01" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M02 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M02" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M03 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M03" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M04 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M04" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M05 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M05" & vbCrLf _
            & ", Sum (Decode(RSTBUDR1.OPS_YYYYPP,'" & SEASON_YEAR & M06 & "',NVL(RSTBUDR1.BUDGET,0),0)) MX_M06" & vbCrLf _
            & " from " & RSTBUDR1 & " RSTBUDR1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTBUDR1.OPS_YYYYPP between '" & SEASON_YEAR & M01 & "' and '" & SEASON_YEAR & M06 & "'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTBUDR1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTBUDR1") _
            & " group by RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ") group by CUST_CODE, CUST_STORE_NO, BRAND_CODE, HC_CODE, COLLECTION_GENDER" & vbCrLf _
            & ") X " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_STATUS = 'A'" _
            & sqlAuths

        SPTMBPL1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Update " & SPTMBPL1 & " Set RB_SSN = NVL(RB_M01,0)+NVL(RB_M02,0)+NVL(RB_M03,0)+NVL(RB_M04,0)+NVL(RB_M05,0)+NVL(RB_M06,0)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Lookups")


        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Customer / HC Summary")

        ws = wb.Worksheets("ChainSummary")

        ASCMAIN1.sql = "Select Distinct " & vbCrLf _
            & "  ICTCOLL0.HC_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & " from " & SPTMBPL1 & " X" & vbCrLf _
            & ",ICTCOLL0,ARTCUST1" & vbCrLf _
            & " where ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 9
        rTotal = DataTable.Select("", "HC_NAME, CUST_NAME").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "HC_NAME,CUST_NAME")
            r += 1
            ws.Range("B" & CStr(r0 + r) & ":C" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 1000 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next


        For i As Integer = 4 To 16 + 3 ' sp added 3 columns
            If i <> 8 Then
                Dim cc3 As Integer = i
                xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc3))
                xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc3), Excel_Cell(r0 + DataTable.Rows.Count, cc3))
                xlSourceRange.Copy(xlDestRange)
            End If
        Next

        xlSourceRange = ws.Range(Excel_Cell(r0 + 1, 1))
        xlDestRange = ws.Range(Excel_Cell(r0 + 1, 1), Excel_Cell(r0 + DataTable.Rows.Count, 1))
        xlSourceRange.Copy(xlDestRange)

        '"Key_2"
        wb.Names.Add("SummaryByChain", "=CHAINPLANNER!$A$" & CStr(r0) & ":$T$" & CStr(r0 + DataTable.Rows.Count))


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
            & ", X.RS_SSN, X.RS_M01, X.RS_M02, X.RS_M03, X.RS_M04, X.RS_M05, X.RS_M06" & vbCrLf _
            & ", X.MX_SSN, X.MX_M01, X.MX_M02, X.MX_M03, X.MX_M04, X.MX_M05, X.MX_M06" & vbCrLf _
            & ", X.RB_SSN, X.RB_M01, X.RB_M02, X.RB_M03, X.RB_M04, X.RB_M05, X.RB_M06" & vbCrLf _
            & " from " & SPTMBPL1 & " X" & vbCrLf _
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

        r0 = 10
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_NAME,HC_NAME,CUST_STORE_NO")
            r += 1
            ws.Range("B" & CStr(r0 + r) & ":O" & CStr(r0 + r)).Value2 = row.ItemArray

            ws.Cells(r0 + r, 16).Value2 = row.Item("MX_SSN")
            ws.Cells(r0 + r, 18).Value2 = row.Item("RS_SSN")
            ws.Cells(r0 + r, 24).Value2 = row.Item("RB_SSN")

            ws.Cells(r0 + r, 30).Value2 = row.Item("RB_M01")
            ws.Cells(r0 + r, 31).Value2 = row.Item("RB_M02")
            ws.Cells(r0 + r, 32).Value2 = row.Item("RB_M03")
            ws.Cells(r0 + r, 33).Value2 = row.Item("RB_M04")
            ws.Cells(r0 + r, 34).Value2 = row.Item("RB_M05")
            ws.Cells(r0 + r, 35).Value2 = row.Item("RB_M06")

            ws.Cells(r0 + r, 36).Value2 = row.Item("MX_M01")
            ws.Cells(r0 + r, 38).Value2 = row.Item("MX_M02")
            ws.Cells(r0 + r, 40).Value2 = row.Item("MX_M03")
            ws.Cells(r0 + r, 42).Value2 = row.Item("MX_M04")
            ws.Cells(r0 + r, 44).Value2 = row.Item("MX_M05")
            ws.Cells(r0 + r, 46).Value2 = row.Item("MX_M06")

            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=MAINDATA!$A$" & CStr(r0) & ":$AA$" & CStr(r0 + DataTable.Rows.Count))

        wb.Names.Add("KEY", "=MAINDATA!$A$" & CStr(r0) & ":$A$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CURR_SN", "=MAINDATA!$R$" & CStr(r0) & ":$R$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CURR_SN_LY", "=MAINDATA!$Q$" & CStr(r0) & ":$Q$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("Fall_14", "=MAINDATA!$P$" & CStr(r0) & ":$P$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("LYFreelance", "=MAINDATA!$P$" & CStr(r0) & ":$P$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("LYSales", "=MAINDATA!$R$" & CStr(r0) & ":$R$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("FC_2015", "=MAINDATA!$X$" & CStr(r0) & ":$X$" & CStr(r0 + DataTable.Rows.Count))

        wb.Names.Add("TotalAllocation", "=MAINDATA!$BD$" & CStr(r0) & ":$BD$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("TYSalesPlan", "=MAINDATA!$X$" & CStr(r0) & ":$X$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CountFreelance", "=MAINDATA!$BE$" & CStr(r0) & ":$BE$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("CountRetail", "=MAINDATA!$BF$" & CStr(r0) & ":$BF$" & CStr(r0 + DataTable.Rows.Count))
        wb.Names.Add("RetailFcstddoors", "=MAINDATA!$BG$" & CStr(r0) & ":$BG$" & CStr(r0 + DataTable.Rows.Count))
        '
        ' ws.Cells(r0 - 2, 15).Value = SEASON_CODE_2LY
        ws.Cells(r0 - 2, 2).Value = SEASON_CODE

        'ws.Cells(r0 - 2, 21).Value = SEASON_CODE_PRIOR_LY & " STD"
        'ws.Cells(r0 - 2, 22).Value = SEASON_CODE_PRIOR & " STD"
        'ws.Cells(r0 - 2, 23).Value = SEASON_CODE_PRIOR & " Trend"
        'ws.Cells(r0 - 2, 24).Value = SEASON_CODE

        'ws.Cells(r0 - 2, 40).Value = SEASON_CODE_LY & " Total"
        'ws.Cells(r0 - 2, 41).Value = SEASON_CODE & " Total"

        ASCMAIN1.Progress("-", "Formulas")

        For i As Integer = 1 To 59
            Dim F As String = ws.Range(Excel_Cell(r0 + 1, i)).Formula ' ws.Cells(r0 + 1, i).formula
            If F.StartsWith("=") Then
                xlSourceRange = ws.Range(Excel_Cell(r0 + 1, i))
                xlDestRange = ws.Range(Excel_Cell(r0 + 1, i), Excel_Cell(r0 + DataTable.Rows.Count, i))
                xlSourceRange.Copy(xlDestRange)
            End If
        Next

        'xlSourceRange = ws.Range("A" & CStr(r0 + 1) & ":A" & CStr(r0 + 1))
        'xlDestRange = ws.Range("A" & CStr(r0 + 1) & ":A" & CStr(r0 + DataTable.Rows.Count))
        'xlSourceRange.Copy(xlDestRange)

        'Dim cc2 As Integer

        'For i As Integer = 23 To 26
        '    If i = 25 Then
        '    Else
        '        cc2 = i
        '        xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc2))
        '        xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc2), Excel_Cell(r0 + DataTable.Rows.Count, cc2))
        '        xlSourceRange.Copy(xlDestRange)
        '    End If
        'Next
         
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
                XLS_FILENAME = "Freelance_Budget_Planner"
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