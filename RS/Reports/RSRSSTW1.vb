
Imports Microsoft.Office.Interop

Public Class RSRSSTW1

    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String
    Dim SEASON_YEAR_LY As String

    Dim RSTSSTW1 As String

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
        SQLW &= SQLA_filter("BRAND_CODE", "ICTCOLL1")
        SQLW &= SQLA_filter("COLLECTION_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("COST_CATGY_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("CUST_CODE", "RSTRETL1")
        SQLW &= SQLA_filter("HC_CODE", "ICTCOLL1")
        SQLW &= SQLA_filter("ITEM_BASIC_PROMO", "ICTITEM1")
        SQLW &= SQLA_filter("ITEM_CODE", "RSTRETL1")
        SQLW &= SQLA_filter("PROD_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")
        SQLW &= SQLA_filter("CHANNEL_CODE", "SOTTCLS1")

        'path of the template
        ASCMAIN1.Progress("Now Creating Workbook")
        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.Folders("SharedRoot") = "C:\SHARE\INT\"
        End If
        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW) '202353

        Dim RYM As String = rowGLTPARM3.Item("YYYYMM") '202401
        SEASON_YEAR = Mid(RYM, 1, 4) '2024
        SEASON_YEAR_LY = Format(Val(SEASON_YEAR) - 1, "0000") '2023

        Dim WKS As Integer = Val(Mid(RYW, 5, 2)) + 4 '5
        If WKS > 52 Then WKS = WKS - 52

        If Mid(RYM, 5, 2) >= "07" Then
            SEASON_TYPE = "F"
            WKS = WKS - 26
        Else
            SEASON_TYPE = "S"
        End If
        SEASON_CODE = SEASON_YEAR & SEASON_TYPE

        Dim M(6) As String
        M(0) = IIf(SEASON_TYPE = "S", SEASON_YEAR_LY & "12", SEASON_YEAR & "06") '202312
        M(1) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "01", SEASON_YEAR & "07") '202401
        M(2) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "02", SEASON_YEAR & "08") '202402
        M(3) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "03", SEASON_YEAR & "09") '202403
        M(4) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "04", SEASON_YEAR & "10") '202404
        M(5) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "05", SEASON_YEAR & "11") '202405
        M(6) = IIf(SEASON_TYPE = "S", SEASON_YEAR & "06", SEASON_YEAR & "12") '202406


        Dim SQLP As String = ""
        Dim SQLP0 As String = ""
        Dim W As Integer = 0
        'previous month (202344-202348)
        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & M(0) & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
            Dim yearWeek As String = row.Item("YYYYWW").ToString()
            Dim weekPart As String = Mid(yearWeek, 5, 2)
            W += 1
            Dim weekColumn As String = "ST_P" & W.ToString("00")
            SQLP0 &= ", 0 " & weekColumn

            SQLP &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & yearWeek & "',RSTRETL1.QTY_SOLD,0)) " & weekColumn & vbCrLf
        Next
        SQLP0 &= vbCrLf

        W = 0
        Dim SQLM As String = ""
        Dim SQLM0 As String = ""
        Dim YWs(27) As String
        For I As Integer = 1 To 6
            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & M(I) & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                Dim yearWeek As String = row.Item("YYYYWW").ToString()
                Dim weekPart As String = Mid(yearWeek, 5, 2)
                W += 1
                Dim weekColumn As String = "ST_W" & W.ToString("00")
                SQLM0 &= ", 0 " & weekColumn
                YWs(W) = yearWeek
                SQLM &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & yearWeek & "',RSTRETL1.QTY_SOLD,0)) " & weekColumn & vbCrLf
            Next
            SQLM0 &= vbCrLf
        Next

        SQLM &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & RYW & "',RSTRETL1.QTY_EOW,0)) ST_ONH" & vbCrLf
        SQLM0 &= ", 0 ST_ONH" & vbCrLf

        Dim sqlAuths As String = IIf(chkALLDOORS.Checked, "", vbCrLf _
            & "   and SATAUTH1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SATAUTH1.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and SATAUTH1.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_CLOSED IS NULL" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL")

        ASCMAIN1.Progress("Work Table")
        Dim wk27 As String = ", 0 ST_WXX"
        If (YWs(27) & "") <> "" Then
            wk27 = ", Sum (ST_W27) ST_W27"
        End If

        ASCMAIN1.sql = "Select X.* from " & IIf(chkALLDOORS.Checked, "", "SATAUTH1, ") & "ICTBRAN1, ICTCOLL0, (" & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, TRADE_CLASS_CODE, BRAND_CODE, HC_CODE, COLLECTION_GENDER, COLLECTION_CODE, PROD_CODE, COST_CATGY_CODE, ITEM_BASIC_PROMO, ITEM_CODE" & vbCrLf _
            & ", Sum (QTY_OPEN) QTY_OPEN" & vbCrLf _
            & ", Sum (SI_TOTAL) SI_TOTAL, Sum (SI_M00) SI_M00, Sum (SI_M01) SI_M01, Sum (SI_M02) SI_M02, Sum (SI_M03) SI_M03, Sum (SI_M04) SI_M04, Sum (SI_M05) SI_M05, Sum (SI_M06) SI_M06" & vbCrLf _
            & ", Sum (ST_TOTAL) ST_TOTAL, Sum (ST_P01) ST_P01, Sum (ST_P02) ST_P02, Sum (ST_P03) ST_P03, Sum (ST_P04) ST_P04, Sum (ST_P05) ST_P05" & vbCrLf _
            & ", Sum (ST_W01) ST_W01, Sum (ST_W02) ST_W02, Sum (ST_W03) ST_W03, Sum (ST_W04) ST_W04" & vbCrLf _
            & ", Sum (ST_W05) ST_W05, Sum (ST_W06) ST_W06, Sum (ST_W07) ST_W07, Sum (ST_W08) ST_W08" & vbCrLf _
            & ", Sum (ST_W09) ST_W09, Sum (ST_W10) ST_W10, Sum (ST_W11) ST_W11, Sum (ST_W12) ST_W12, Sum (ST_W13) ST_W13" & vbCrLf _
            & ", Sum (ST_W14) ST_W14, Sum (ST_W15) ST_W15, Sum (ST_W16) ST_W16, Sum (ST_W17) ST_W17" & vbCrLf _
            & ", Sum (ST_W18) ST_W18, Sum (ST_W19) ST_W19, Sum (ST_W20) ST_W20, Sum (ST_W21) ST_W21" & vbCrLf _
            & ", Sum (ST_W22) ST_W22, Sum (ST_W23) ST_W23, Sum (ST_W24) ST_W24, Sum (ST_W25) ST_W25, Sum (ST_W26) ST_W26" & wk27 & vbCrLf _
            & ", Sum (ST_ONH) ST_ONH" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, RSTRETL1.ITEM_CODE" & vbCrLf _
            & ", 0 QTY_OPEN" & vbCrLf _
            & ", 0 SI_TOTAL, 0 SI_M00, 0 SI_M01, 0 SI_M02, 0 SI_M03, 0 SI_M04, 0 SI_M05, 0 SI_M06" & vbCrLf _
            & ", Sum (RSTRETL1.QTY_SOLD) ST_TOTAL" & SQLP & SQLM _
            & " from RSTRETL1, ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP between '" & M(0) & "' and '" & M(6) & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW <= '" & RYW & "'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTRETL1") _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, RSTRETL1.ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", 0 QTY_OPEN" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) SI_TOTAL" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(0) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M00" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(1) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M01" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(2) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M02" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(3) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M03" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(4) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M04" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(5) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M05" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & M(6) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SI_M06" & vbCrLf _
            & ", 0 ST_TOTAL" & SQLP0 & SQLM0 _
            & " from SOTINVH2, ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & M(0) & "' and '" & M(6) & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and SOTINVH2.OPS_YYYYWW <= '" & RYW & "'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "SOTINVH2") _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') BETWEEN '" & M(0) & "' AND '" & M(6) & "' THEN SOTORDR2.ORDR_QTY_OPEN ELSE 0 END) QTY_OPEN" & vbCrLf _
            & ", 0 SI_TOTAL, 0 SI_M00, 0 SI_M01, 0 SI_M02, 0 SI_M03, 0 SI_M04, 0 SI_M05, 0 SI_M06" & vbCrLf _
            & ", 0 ST_TOTAL" & SQLP0 & SQLM0 _
            & " from SOTORDR1, SOTORDR2, ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1" & vbCrLf _
            & " where SOTORDR2.ORDR_STATUS between 'O' and 'P'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "SOTORDR2") _
            & " group by SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, ARTCUST1.TRADE_CLASS_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ") group by CUST_CODE, CUST_STORE_NO, TRADE_CLASS_CODE, BRAND_CODE, HC_CODE, COLLECTION_GENDER, COLLECTION_CODE, PROD_CODE, COST_CATGY_CODE, ITEM_BASIC_PROMO, ITEM_CODE" & vbCrLf _
            & ") X " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_STATUS = 'A'" _
            & sqlAuths
        RSTSSTW1 = ASCMAIN1.Temp_Table


        'ASCMAIN1.sql = "Update " & RSTSSTW1 & " Set RB_SSN = NVL(RB_M01,0)+NVL(RB_M02,0)+NVL(RB_M03,0)+NVL(RB_M04,0)+NVL(RB_M05,0)+NVL(RB_M06,0)"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Lookups")

        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Items")

        ws = wb.Worksheets("Lookup")

        ASCMAIN1.sql = "Select Distinct " & vbCrLf _
            & "X.ITEM_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTCOLL1.COLLECTION_GENDER, ICTITEM1.PROD_CODE" & vbCrLf _
            & " from (Select Distinct ITEM_CODE from " & RSTSSTW1 & ") X" & vbCrLf _
            & ",ICTCOLL1,ICTITEM1,ICTBRAN1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 1 ' Row containing headings
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "SALES_DIVISION_CODE,BRAND_CODE,HC_CODE,ITEM_CODE")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":G" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 1000 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next

        'For i As Integer = 4 To 16
        '    If i <> 8 Then
        '        Dim cc3 As Integer = i
        '        xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc3))
        '        xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc3), Excel_Cell(r0 + DataTable.Rows.Count, cc3))
        '        xlSourceRange.Copy(xlDestRange)
        '    End If
        'Next

        'xlSourceRange = ws.Range(Excel_Cell(r0 + 1, 1))
        'xlDestRange = ws.Range(Excel_Cell(r0 + 1, 1), Excel_Cell(r0 + DataTable.Rows.Count, 1))
        'xlSourceRange.Copy(xlDestRange)

        wb.Names.Add("lookup_brand", "=Lookup!$A$" & CStr(r0) & ":$G$" & CStr(r0 + DataTable.Rows.Count))

        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("Data")

        Dim wxx As String = ""
        Dim w27 As String = ""
        If (YWs(27) & "" = "") Then ' 26 weeks
            wxx = ", 0 WXX"
        Else ' 27 weeks
            w27 = ", X.ST_W27"
        End If

        ASCMAIN1.sql = "Select " & vbCrLf _
        & " TO_CHAR(SYSDATE,'MM/DD/YYYY') DATEOFDATA" & vbCrLf _
        & ", SOTSELL1.REGION_CODE ASDNBR" & vbCrLf _
        & ", SOTSREG1.REGION_DESC ASDNAME" & vbCrLf _
        & ", ARTCUST2.SELL_CODE AENBR" & vbCrLf _
        & ", SOTSELL1.SELL_NAME AEName" & vbCrLf _
        & ", SOTSELL1_AC.SELL_NAME AC" & vbCrLf _
        & ", ICTITEM1.PROD_CODE ItemType" & vbCrLf _
        & ", ICTCOLL1.HC_CODE BrNum" & vbCrLf _
        & ", ICTCOLL0.HC_NAME BrandName" & vbCrLf _
        & ", ICTITEM1.COLLECTION_CODE SubBrand" & vbCrLf _
        & ", ICTCOLL1.COLLECTION_NAME SubBrandName" & vbCrLf _
        & ", ARTCUST1.CUST_NAME ChainName" & vbCrLf _
        & ", ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO DoorID" & vbCrLf _
        & ", ARTCUST2.CUST_STORE_NAME CUSTOMER_Name" & vbCrLf _
        & ", ICTITEM1.ITEM_DESC DESCRIPTION_COMMENT" & vbCrLf _
        & ", ICTITEM1.ITEM_RETAIL_PRICE LIST_PRICE_PRICING_UNIT" & vbCrLf _
        & ", ICTCOLL1.BRAND_CODE BRAND" & vbCrLf _
        & ", X.ITEM_CODE ITEM_#" & vbCrLf _
        & ", X.CUST_CODE CUST#_PART1" & vbCrLf _
        & ", X.CUST_STORE_NO CUST#_PART2" & vbCrLf _
        & ", X.QTY_OPEN" & vbCrLf _
        & ", X.SI_TOTAL, X.SI_M00, X.SI_M01, X.SI_M02, X.SI_M03, X.SI_M04, X.SI_M05, X.SI_M06" & vbCrLf _
        & ", X.ST_TOTAL, X.ST_P01, X.ST_P02, X.ST_P03, X.ST_P04, X.ST_P05" & vbCrLf _
        & $", X.ST_W01, X.ST_W02, X.ST_W03, X.ST_W04{wxx}" & vbCrLf _
        & ", X.ST_W05, X.ST_W06, X.ST_W07, X.ST_W08" & vbCrLf _
        & ", X.ST_W09, X.ST_W10, X.ST_W11, X.ST_W12, X.ST_W13" & vbCrLf _
        & ", X.ST_W14, X.ST_W15, X.ST_W16, X.ST_W17" & vbCrLf _
        & ", X.ST_W18, X.ST_W19, X.ST_W20, X.ST_W21" & vbCrLf _
        & $", X.ST_W22, X.ST_W23, X.ST_W24, X.ST_W25, X.ST_W26{w27}" & vbCrLf _
        & ", X.ST_ONH" & vbCrLf _
        & " from " & RSTSSTW1 & " X" & vbCrLf _
        & ",ICTBRAN1,ARTCUST1,ARTCUST2,SOTSELL1,SOTSREG1,ICTCOLL0,ICTCOLL1,ICTITEM1,SOTSELL1 SOTSELL1_AC" & vbCrLf _
        & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
        & "   and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
        & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
        & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
        & "   and ICTCOLL0.HC_CODE (+) = ICTCOLL1.HC_CODE" & vbCrLf _
        & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
        & "   and SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC" & vbCrLf _
        & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
        & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
        & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"


        DataTable = ASCDATA1.GetDataTable

        r0 = 5 ' Row containing headings
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "ChainName,BRAND,BrandName") ' "CUST_NAME,HC_NAME,CUST_STORE_NO")
            r += 1
            ws.Range("E" & CStr(r0 + r) & ":BO" & CStr(r0 + r)).Value2 = row.ItemArray

            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("Pivot_Base", "=Data!$A$" & CStr(r0) & ":$BO$" & CStr(r0 + DataTable.Rows.Count))

        ws.Cells(1, 1).Value = Now
        ws.Cells(2, 1).Value = SEASON_CODE
        ws.Cells(3, 1).Value = rowGLTPARM3.Item("LEGEND")
        Dim YEAR_AND_SEASON As String = ws.Cells(2, 1).Value.ToString()
        Dim YEAR As String = YEAR_AND_SEASON.Substring(0, 4)
        Dim SEASON As String = YEAR_AND_SEASON.Substring(4, 1)
        Dim SEASON_NAME As String = ""
        If SEASON = "S" Then
            SEASON_NAME = "Spring"
        ElseIf SEASON = "F" Then
            SEASON_NAME = "Fall"
        End If
        Dim SPRING_MONTHS As String() = {"Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun"}
        Dim FALL_MONTHS As String() = {"Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        Dim SELECTED_MONTHS As String() = If(SEASON_TYPE = "S", SPRING_MONTHS, FALL_MONTHS)
        Dim sellThruByWeekSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("Sellthru by Week")
        Dim sellThruByMonthSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("Sellthru by Month")

        ws.Cells(4, 1).Value = SEASON_NAME & " " & YEAR & " Sell Thru" 'data A4
        Dim seasonAndYear As String = ws.Cells(4, 1).Value.ToString().Substring(0, ws.Cells(4, 1).Value.ToString().IndexOf(" Sell Thru"))
        sellThruByMonthSheet.Cells(6, 7).Value = seasonAndYear ' Cell G6

        'In sellthru by month updating headings for L7 thru R7 
        For i As Integer = 0 To 6
            sellThruByMonthSheet.Cells(7, i + 12).Value = SELECTED_MONTHS(i) & " ST"
        Next

        Dim currentMonth As String = String.Empty
        Dim weekNumber As Integer = 0
        Dim startColumn As Integer = 12 ' Column L corresponds to 12
        Dim col As Integer = 0
        Dim pivotTable As Excel.PivotTable = sellThruByWeekSheet.PivotTables("PivotTable1")
        Dim HEADER_LOG As New List(Of String)

        For MM As Integer = 0 To 6
            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & M(MM) & "' order by YYYYWW"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                Dim YYYYWW As String = row.Item("YYYYWW")
                Dim LEGEND As String = row.Item("LEGEND")
                ' Extract month and week number from LEGEND (e.g., "2023-49 (Jan:1/5)")
                Dim legendParts As String() = LEGEND.Split(New Char() {"(", ":", "/", ")"}, StringSplitOptions.RemoveEmptyEntries)
                If legendParts.Length >= 4 Then
                    Dim monthFromLegend As String = legendParts(1).Trim() 'e.g. Jan
                    Dim weekFromLegend As Integer = Integer.Parse(legendParts(2)) 'e.g. 1
                    If currentMonth <> monthFromLegend Then
                        currentMonth = monthFromLegend
                        weekNumber = 1
                    Else
                        weekNumber += 1
                    End If
                    Dim headerText As String = $"{currentMonth} Wk {weekNumber} Sell Thru"
                    sellThruByWeekSheet.Cells(7, startColumn).Value = headerText
                    HEADER_LOG.Add("Col " & startColumn.ToString() & ": " & headerText)
                    startColumn += 1
                End If
                col += 1
                Dim d As Date = row.Item("WEEK_END_DATE")
                ws.Cells(2, 34 + col).Value = Format(d, "yyMMdd") & "Q"
                ws.Cells(3, 34 + col).Value = Mid(LEGEND, 10, 7)
            Next
            If MM = 1 And weekNumber = 4 Then
                col += 1

                Dim xCol As Integer = startColumn ' <-- the column you're about to write "X" into
                sellThruByWeekSheet.Cells(7, xCol).Value = "X"

                ' Hide that pivot column
                sellThruByWeekSheet.Columns(xCol).Hidden = True

                startColumn += 1
            End If
        Next
        ws.Cells(1, 35 + col).Value = "QTYOH"

        For I As Integer = 0 To 6
            ws.Cells(2, 27 + I).Value = "SI" & Mid(M(I), 3, 4) & " Q"
            Dim D As Date = Mid(M(I), 5, 2) & "/01/" & Mid(M(I), 1, 4)
            ws.Cells(3, 27 + I).Value = Format(D, "MMM") & " SELLIN"
        Next

        '        If Mid(YYYYWW, 5, 2) = "53" Then
        '            ' skip - week 53 combined with week 52
        '        Else
        ASCMAIN1.Progress("-", "Formulas")

        For i As Integer = 1 To 4
            Dim F As String = ws.Range(Excel_Cell(r0 + 1, i)).Formula ' ws.Cells(r0 + 1, i).formula
            If F.StartsWith("=") Then
                xlSourceRange = ws.Range(Excel_Cell(r0 + 1, i))
                xlDestRange = ws.Range(Excel_Cell(r0 + 1, i), Excel_Cell(r0 + DataTable.Rows.Count, i))
                xlSourceRange.Copy(xlDestRange)
            End If
        Next
        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        'pivotTable.RefreshTable()

        For Each pf As Excel.PivotField In pivotTable.DataFields
            If pf.Name.Contains("Wk") AndAlso Not HEADER_LOG.Any(Function(h) pf.Name.Trim().Equals(h.Substring(h.IndexOf(":") + 2).Trim(), StringComparison.OrdinalIgnoreCase)) Then
                pf.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlHidden
            End If
        Next


        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Sell_in_Sell_thru_By_Week"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)

                'wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                '          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
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