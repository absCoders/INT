Public Class RSRFLSH1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 100, -1)
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

        MyBase.Get_SQL("*")

        Dim LYX As Integer = 1
        LYX = Val(cbeLY.Value & "")

        Dim sql_Filter As String = ""

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Dim rg As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        ASCMAIN1.Progress("-", "Lookups")

        ws = wb.Worksheets("Info & Lookups")

        ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ICTCOLL0.HC_NAME, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE Frch_Abbrev, 'IPLB' BUS_UNIT_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.HC_CODE, 'No Bonus' FC" & vbCrLf _
            & " from ICTCOLL1,ICTBRAN1,ICTCOLL0 " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE (+) = ICTCOLL1.HC_CODE"
        DataTable = ASCDATA1.GetDataTable

        r = 0
        For Each row As DataRow In DataTable.Select("", "COLLECTION_CODE")
            r += 1
            ws.Range("F" & CStr(4 + r) & ":O" & CStr(4 + r)).Value2 = row.ItemArray
        Next
        If ASCMAIN1.CLIENT = "INT" Then
            wb.Names.Add("BrandList", "='Info & Lookups'!$F$4:$O$" & CStr(4 + DataTable.Rows.Count))
        Else
            wb.Names.Add("Collections", "='Info & Lookups'!$F$4:$O$" & CStr(4 + DataTable.Rows.Count))
        End If


        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE, 'Monday' EDI_GROUP" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & " union Select NULL CUST_CODE, '*' CUST_NAME, NULL TRADE_CLASS_CODE, NULL EDI_GROUP from Dual"

        DataTable = ASCDATA1.GetDataTable

        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_CODE")
            r += 1
            ws.Range("A" & CStr(4 + r) & ":D" & CStr(4 + r)).Value2 = row.ItemArray
        Next
        If ASCMAIN1.CLIENT = "INT" Then
            wb.Names.Add("ChainList", "='Info & Lookups'!$A$4:$D$" & CStr(4 + DataTable.Rows.Count))
        Else
            wb.Names.Add("Customers", "='Info & Lookups'!$A$4:$D$" & CStr(4 + DataTable.Rows.Count))
        End If


        If ASCMAIN1.CLIENT = "INT" Then
            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO CUSTSTORE, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", SOTSDSC1.SDS_NAME" & vbCrLf _
                & " from ARTCUST2,SOTSDSC1 where SOTSDSC1.SDS_CODE = ARTCUST2.SDS_CODE"
            DataTable = ASCDATA1.GetDataTable

            r = 0
            For Each row As DataRow In DataTable.Select("", "CUSTSTORE")
                r += 1
                ws.Range("T" & CStr(4 + r) & ":V" & CStr(4 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("AC_Terrs", "='Info & Lookups'!$T$4:$V$" & CStr(4 + DataTable.Rows.Count))
        Else
            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO CUST_STORE" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CUST_RANK_W, ARTCUST2.CUST_STORE_CUST_RANK_M" & vbCrLf _
                & ", ARTCUST2.SELL_CODE, ARTCUST2.CUST_STORE_NAME, SOTSDSC1.SDS_NAME" & vbCrLf _
                & " from ARTCUST2,SOTSDSC1 where SOTSDSC1.SDS_CODE (+) = ARTCUST2.SDS_CODE"
            DataTable = ASCDATA1.GetDataTable

            r = 0
            For Each row As DataRow In DataTable.Select("", "CUST_STORE")
                r += 1
                ws.Range("Q" & CStr(4 + r) & ":V" & CStr(4 + r)).Value2 = row.ItemArray
            Next
            wb.Names.Add("Stores", "='Info & Lookups'!$Q$4:$V$" & CStr(4 + DataTable.Rows.Count))
        End If



        ASCMAIN1.sql = "Select SOTSELL1.SELL_CODE, SOTSELL1.SELL_NAME" & vbCrLf _
            & ", SOTSELL1.REGION_CODE, SOTSREG1.REGION_DESC, SOTSREG1.VP_CODE" & vbCrLf _
            & " from SOTSELL1,SOTSREG1" & vbCrLf _
            & " where SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"
        DataTable = ASCDATA1.GetDataTable

        r = 0
        For Each row As DataRow In DataTable.Select("", "SELL_CODE")
            r += 1
            ws.Range("X" & CStr(4 + r) & ":AB" & CStr(4 + r)).Value2 = row.ItemArray
        Next
        If ASCMAIN1.CLIENT = "INT" Then
            wb.Names.Add("TerrListing", "='Info & Lookups'!$X$4:$AB$" & CStr(4 + DataTable.Rows.Count))
        Else
            wb.Names.Add("Field", "='Info & Lookups'!$X$4:$AB$" & CStr(4 + DataTable.Rows.Count))
        End If




        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("Data")

        Dim WKS(10, 1)
        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")
        Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        Dim LEGEND As String = rowGLTPARM3.Item("LEGEND")



        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & RYM & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
            Dim RW As Integer = Val(row.Item("REL_WEEK") & "")
            Dim YW As String = row.Item("YYYYWW")
            WKS(RW, 0) = YW
            WKS(RW, 1) = ASCMAIN1.Week_Calc(YW, -52 * LYX)
            If ASCMAIN1.CLIENT = "INT" Then
                If chkUSE52WEEKS.Checked Then
                    ' LEAVE LY ALONE
                Else
                    Dim LYW As String = Format(Val(Mid(YW, 1, 4)) - LYX, "0000") & Mid(YW, 5, 2)
                    WKS(RW, 1) = LYW
                End If
            End If
        Next

        Dim WKS_ADJ As Integer = 0
        If ASCMAIN1.CLIENT = "INT" Then
            'When we ran the retail flash this morning, last year’s retails for Jan wk. 5 are not pulling.  These should be visible on the month by week tab.
            'cm thinks that Jan 19 (a 4 week month) should be compared to the 5 weeks in Jan'18
            If WKS(4, 1) = "201752" Then
                WKS(5, 1) = "201753"
            End If
            ' PROBABLY NEED TO ADDRESS THIS AGAIN IN 2023 OR 2024
            'If WKS(4, 1) = "202352" Then
            '    WKS(5, 1) = "202353"
            'End If

            Dim YW As String = WKS(4, 1) ' ENDING WEEK FOR LY
            Dim WW As String = Mid(YW, 5, 2)
            If WW = "52" Then
                Dim YYYY As String = Mid(YW, 1, 4)
                ASCMAIN1.sql = $"SELECT COUNT (*) WEEKS FROM GLTPARM3 WHERE YYYYWW LIKE '{YYYY}%'"
                Dim WEEKS_IN_YYYY As Integer = Val(ASCDATA1.GetDataValue & "")
                If WEEKS_IN_YYYY = 53 Then
                    WKS_ADJ = 1
                    WKS(5, 1) = YYYY & "53"
                End If
            End If
        End If

        If Mid(RYM, 5, 2) >= "07" Then
            WKS(7, 0) = Mid(RYW, 1, 4) & "23"  'HTD445
        Else
            WKS(7, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -4) 'HTD445
            ASCMAIN1.sql = "Select Min (YYYYWW) YYYYWW from GLTPARM3 where YYYYMM = '" & Mid(RYM, 1, 4) & "01" & "'"
            Dim row As DataRow = ASCDATA1.GetDataRow
            WKS(7, 0) = row.Item(0)
        End If
        If Mid(RYW, 5, 2) >= "27" Then
            WKS(8, 0) = Mid(RYW, 1, 4) & "27" 'STD454
        Else
            WKS(8, 0) = Mid(RYW, 1, 4) & "01" 'STD454
        End If
        If Mid(RYM, 5, 2) = "01" Then
            WKS(9, 0) = ASCMAIN1.Week_Calc(RYW, 1 - REL_WEEK) 'YTD445
        Else


            Dim WEEKS_OFFSET As Integer = 4

            'If Mid(RYM, 5, 2) = "02" Then
            ASCMAIN1.sql = "Select COUNT (*) from GLTPARM3 where YYYYWW LIKE '" & Format(Val(Mid(RYM, 1, 4)) - LYX, "0000") & "%'"
            Dim WEEKS_LY As Integer = Val(ASCDATA1.GetDataValue)
            If WEEKS_LY = 53 Then
                WEEKS_OFFSET = 5
            End If
            'End If

            WKS(9, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -1 * WEEKS_OFFSET) 'YTD445
        End If

        WKS(10, 0) = Mid(RYW, 1, 4) & "01" 'YTD454

        For W As Integer = 7 To 10
            WKS(W, 1) = ASCMAIN1.Week_Calc(WKS(W, 0), -52)
            If ASCMAIN1.CLIENT = "INT" Then
                If chkUSE52WEEKS.Checked Then
                    ' LEAVE LY ALONE
                Else
                    Dim YW = WKS(W, 0)
                    Dim LYW As String = Format(Val(Mid(YW, 1, 4)) - LYX, "0000") & Mid(YW, 5, 2)
                    WKS(W, 1) = LYW
                End If
            End If
        Next

        '& ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(6, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK5" & vbCrLf _
        '& ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(6, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK5" & vbCrLf _


        '201544' TY_WK1
        '201444' LY_WK1
        '201545' TY_WK2
        '201445' LY_WK2
        '201546' TY_WK3
        '201446' LY_WK3
        '201547' TY_WK4
        '201447' LY_WK4
        '201548' TY_WK5
        '201448' LY_WK5
        'between '201544' and '201545' TY_MTD
        'between '201444' and '201445' LY_MTD
        'between '201523' and '201545' TY_HTD445
        'between '201423' and '201445' LY_HTD445
        'between '201527' and '201545' TY_STD454
        'between '201427' and '201445' LY_STD454
        'between '201449' and '201545' TY_YTD445
        'between '201349' and '201445' LY_YTD445
        'between '201501' and '201545' TY_YTD454
        'between '201401' and '201445' LY_YTD454
        '& " where RSTRETL1.OPS_YYYYWW between '201349' and '201545'" _

        '  sql_WHERE = Replace(sql_WHERE, "RSTCONP1.CUST_CODE || '-' || RSTCONP1.CUST_STORE_NO", "RSTRETL1.CUST_CODE || '-' || RSTRETL1.CUST_STORE_NO")



        '***********
        '        SELECT CUST_CODE, CUST_STORE_NO, COLLECTION_CODE
        ', Sum (TY_WK1) TY_WK1
        ', Sum (LY_WK1) LY_WK1
        ', Sum (TY_WK1) TY_WK2
        ', Sum (LY_WK1) LY_WK2
        ', Sum (TY_WK1) TY_WK3
        ', Sum (LY_WK1) LY_WK3
        ', Sum (TY_WK1) TY_WK4
        ', Sum (LY_WK1) LY_WK4
        ', Sum (TY_WK1) TY_WK5
        ', Sum (LY_WK1) LY_WK5 FROM (
        'Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201805',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK1
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201705',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK1
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201806',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK2
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201706',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK2
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201807',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK3
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201707',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK3
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201808',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK4
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201708',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK4
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201809',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK5
        ', Sum (Decode(RSTRETL1.OPS_YYYYWW,'201709',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK5
        ' from RSTRETL1,ICTITEM1
        ' where RSTRETL1.OPS_YYYYWW between '201649' and '201809'
        ' AND ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE
        'AND RSTRETL1.CUST_CODE = 'BERGDORF'
        ' group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE
        'UNION
        'Select DISTINCT SATAUTH2.CUST_CODE, SATAUTH2.CUST_STORE_NO, ICTCOLL1.COLLECTION_CODE
        ', 0 TY_WK1
        ', 0 LY_WK1
        ', 0 TY_WK2
        ', 0 LY_WK2
        ', 0 TY_WK3
        ', 0 LY_WK3
        ', 0 TY_WK4
        ', 0 LY_WK4
        ', 0 TY_WK5
        ', 0 LY_WK5
        'FROM 
        'SATAUTH2, ICTCOLL1 
        'WHERE
        'SATAUTH2.HC_CODE = ICTCOLL1.HC_CODE
        'AND SATAUTH2.CUST_CODE = 'BERGDORF'

        ')
        'GROUP BY CUST_CODE, CUST_STORE_NO, COLLECTION_CODE

        Dim retlSqlSub As String = "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(1, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK1" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(1, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK1" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(2, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK2" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(2, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK2" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(3, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK3" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(3, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK3" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(4, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK4" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(4, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK4" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(5, 0) & "',RSTRETL1.AMT_SOLD / 1000,0)) TY_WK5" & vbCrLf _
            & ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & WKS(5, 1) & "',RSTRETL1.AMT_SOLD / 1000,0)) LY_WK5" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(1, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) TY_MTD" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(1, 1) & "' and '" & WKS(REL_WEEK + WKS_ADJ, 1) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) LY_MTD" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(7, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) TY_HTD445" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(7, 1) & "' and '" & WKS(REL_WEEK + WKS_ADJ, 1) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) LY_HTD445" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(8, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) TY_STD454" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(8, 1) & "' and '" & WKS(REL_WEEK + WKS_ADJ, 1) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) LY_STD454" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(9, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) TY_YTD445" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(9, 1) & "' and '" & WKS(REL_WEEK + WKS_ADJ, 1) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) LY_YTD445" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(10, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) TY_YTD454" & vbCrLf _
            & ", Sum (Case when RSTRETL1.OPS_YYYYWW between '" & WKS(10, 1) & "' and '" & WKS(REL_WEEK + WKS_ADJ, 1) & "' THEN RSTRETL1.AMT_SOLD / 1000 ELSE 0 END) LY_YTD454" & vbCrLf _
            & " from RSTRETL1" & sql_TABLE_NAMEs & vbCrLf _
            & " where RSTRETL1.OPS_YYYYWW between '" & IIf(WKS(9, 1) < WKS(10, 1), WKS(9, 1), WKS(10, 1)) & "' and '" & WKS(REL_WEEK, 0) & "'" & vbCrLf _
            & sql_WHERE & sql_JOIN & sql_Filter & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COLLECTION_CODE"

        Dim RSTRETL1_temp As String = ASCMAIN1.Temp_Table(retlSqlSub)

        ASCMAIN1.sql = "Create Index I_" & RSTRETL1_temp & "_1 on " & RSTRETL1_temp & " (CUST_CODE)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Create Index I_" & RSTRETL1_temp & "_2 on " & RSTRETL1_temp & " (CUST_CODE, CUST_STORE_NO)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Create Index I_" & RSTRETL1_temp & "_3 on " & RSTRETL1_temp & " (CUST_CODE, CUST_STORE_NO, COLLECTION_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.AnalyzeTable(RSTRETL1_temp)

        Dim authSqlSub As String = "Select DISTINCT SATAUTH2.CUST_CODE, SATAUTH2.CUST_STORE_NO, RSTRETL1.ITEM_BASIC_PROMO, ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & ", 0 TY_WK1" & vbCrLf _
            & ", 0 LY_WK1" & vbCrLf _
            & ", 0 TY_WK2" & vbCrLf _
            & ", 0 LY_WK2" & vbCrLf _
            & ", 0 TY_WK3" & vbCrLf _
            & ", 0 LY_WK3" & vbCrLf _
            & ", 0 TY_WK4" & vbCrLf _
            & ", 0 LY_WK4" & vbCrLf _
            & ", 0 TY_WK5" & vbCrLf _
            & ", 0 LY_WK5" & vbCrLf _
            & ", 0 TY_MTD" & vbCrLf _
            & ", 0 LY_MTD" & vbCrLf _
            & ", 0 TY_HTD445" & vbCrLf _
            & ", 0 LY_HTD445" & vbCrLf _
            & ", 0 TY_STD454" & vbCrLf _
            & ", 0 LY_STD454" & vbCrLf _
            & ", 0 TY_YTD445" & vbCrLf _
            & ", 0 LY_YTD445" & vbCrLf _
            & ", 0 TY_YTD454" & vbCrLf _
            & ", 0 LY_YTD454" & vbCrLf _
            & " from SATAUTH2, ICTCOLL1, " & RSTRETL1_temp & " RSTRETL1 " & vbCrLf _
            & " where SATAUTH2.HC_CODE = ICTCOLL1.HC_CODE " & vbCrLf _
            & " and SATAUTH2.CUST_CODE = RSTRETL1.CUST_CODE " & vbCrLf _
                & " and SATAUTH2.CUST_STORE_NO = RSTRETL1.CUST_STORE_NO " & vbCrLf _
            & " and ICTCOLL1.COLLECTION_CODE = RSTRETL1.COLLECTION_CODE "


        Dim sumSqlSub As String = "SELECT CUST_CODE, CUST_STORE_NO, ITEM_BASIC_PROMO, COLLECTION_CODE" _
            & " , Sum (TY_WK1) TY_WK1" _
            & " , Sum (LY_WK1) LY_WK1" _
            & " , Sum (TY_WK2) TY_WK2" _
            & " , Sum (LY_WK2) LY_WK2" _
            & " , Sum (TY_WK3) TY_WK3" _
            & " , Sum (LY_WK3) LY_WK3" _
            & " , Sum (TY_WK4) TY_WK4" _
            & " , Sum (LY_WK4) LY_WK4" _
            & " , Sum (TY_WK5) TY_WK5" _
            & " , Sum (LY_WK5) LY_WK5" _
            & ", Sum (TY_MTD) TY_MTD" & vbCrLf _
            & ", Sum (LY_MTD) LY_MTD" & vbCrLf _
            & ", Sum (TY_HTD445) TY_HTD445" & vbCrLf _
            & ", Sum (LY_HTD445) LY_HTD445" & vbCrLf _
            & ", Sum (TY_STD454) TY_STD454" & vbCrLf _
            & ", Sum (LY_STD454) LY_STD454" & vbCrLf _
            & ", Sum (TY_YTD445) TY_YTD445" & vbCrLf _
            & ", Sum (LY_YTD445) LY_YTD445" & vbCrLf _
            & ", Sum (TY_YTD454) TY_YTD454" & vbCrLf _
            & ", Sum (LY_YTD454) LY_YTD454" & vbCrLf _
            & " FROM (Select * from " & RSTRETL1_temp & vbCrLf & " union " & vbCrLf & authSqlSub & " ) " & vbCrLf _
            & " GROUP BY CUST_CODE, CUST_STORE_NO, ITEM_BASIC_PROMO, COLLECTION_CODE"

        Dim flashSql As String = "Select " & vbCrLf _
            & "  GLTPARM3.LEGEND MONTH" & vbCrLf _
            & ", GLTPARM3.REL_WEEK WKNUM" & vbCrLf _
            & ", X.CUST_CODE || '-' || X.CUST_STORE_NO CUS1CUS2" & vbCrLf _
            & ", X.CUST_CODE CUST1" & vbCrLf _
            & ", ARTCUST1.CUST_NAME CHNAME" & vbCrLf _
            & ", X.CUST_STORE_NO CUST2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DRNAME" & vbCrLf _
            & ", ARTCUST2.SELL_CODE AENUM" & vbCrLf _
            & ", SOTSELL1.SELL_NAME AE_NAME" & vbCrLf _
            & ", SOTSELL1.REGION_CODE ASDNUM" & vbCrLf _
            & ", SOTSREG1.REGION_DESC ASD_NAME" & vbCrLf _
            & ", SOTSELL1_AC.SELL_NAME AC" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", 'Brand' REC_TYPE" & vbCrLf _
            & ", X.COLLECTION_CODE LINE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_NAME LINE_NAME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE CORP" & vbCrLf _
            & ", X.ITEM_BASIC_PROMO BP" & vbCrLf _
            & ", X.TY_WK1" & vbCrLf _
            & ", X.LY_WK1" & vbCrLf _
            & ", X.TY_WK2" & vbCrLf _
            & ", X.LY_WK2" & vbCrLf _
            & ", X.TY_WK3" & vbCrLf _
            & ", X.LY_WK3" & vbCrLf _
            & ", X.TY_WK4" & vbCrLf _
            & ", X.LY_WK4" & vbCrLf _
            & ", X.TY_WK5" & vbCrLf _
            & ", X.LY_WK5" & vbCrLf _
            & ", X.TY_MTD" & vbCrLf _
            & ", X.LY_MTD" & vbCrLf _
            & ", X.TY_HTD445" & vbCrLf _
            & ", X.LY_HTD445" & vbCrLf _
            & ", X.TY_STD454" & vbCrLf _
            & ", X.LY_STD454" & vbCrLf _
            & ", X.TY_YTD445" & vbCrLf _
            & ", X.LY_YTD445" & vbCrLf _
            & ", X.TY_YTD454" & vbCrLf _
            & ", X.LY_YTD454" & vbCrLf _
            & ", ICTCOLL1.HC_CODE HC_CODE" & vbCrLf _
            & " from (" & sumSqlSub & ") X" & vbCrLf _
            & ",ICTCOLL1,ICTBRAN1,ARTCUST1,ARTCUST2,GLTPARM3,SOTSELL1,SOTSREG1,SOTSELL1 SOTSELL1_AC" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   and SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
            & "   and GLTPARM3.YYYYWW = '" & RYW & "'"
        '            & "   and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _

        'ASCMAIN1.sql = "Select R.*, CASE WHEN S.OPS_YYYYPP_OPENED IS NOT NULL THEN CASE WHEN S.OPS_YYYYPP_CLOSED IS NULL THEN 'O' ELSE 'C' END ELSE 'C' END CLOSED_DOORS" _
        '            & ", NULL CHAIN_REGION" & vbCrLf _
        '            & ", NULL CHAIN_DISTRICT" & vbCrLf _
        '            & " FROM ( " & flashSql & " ) R " _
        '            & ", SATAUTH1 S " _
        '            & " WHERE S.CUST_CODE (+) = R.CUST1 " _
        '            & " AND S.CUST_STORE_NO (+) = R.CUST2 " _
        '            & " AND S.HC_CODE (+) = R.HC_CODE "

        ASCMAIN1.sql = "Select R.*, CASE WHEN NVL(A.CUST_STORE_STATUS,'I') = 'A' THEN 'O' ELSE 'C' END CLOSED_DOORS" _
                    & ", NULL CHAIN_REGION" & vbCrLf _
                    & ", NULL CHAIN_DISTRICT" & vbCrLf _
                    & " FROM ( " & flashSql & " ) R " _
                    & ", SATAUTH1 S " _
                    & ", ARTCUST2 A " _
                    & " WHERE S.CUST_CODE (+) = R.CUST1 " _
                    & " AND S.CUST_STORE_NO (+) = R.CUST2 " _
                    & " AND S.HC_CODE (+) = R.HC_CODE " _
                    & " AND A.CUST_CODE (+) = R.CUST1 " _
                    & " AND A.CUST_STORE_NO (+) = R.CUST2 "

        'from WJZ to Dani Tuesday, October 3, 2023 3:43 PM
        'On the flash
        'Change
        '        Show door if it Is open in the brand/collections reported
        'To
        'Show door if it Is Active in the Store Master Or if it Is open I the brand/collections reported
        'Make a list of Stores that are Not Active in the Store Master yet are open (still) in the Store / Brand matrix


        DataTable = ASCDATA1.GetDataTable



        Dim ssgx As String = ASCMAIN1.Folders("Temp") & XNO & ".xlsX"

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = worksheet.Cells("A1")
        range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
        workbook.SaveAs(ssgx, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        range = Nothing
        worksheet = Nothing
        workbook = Nothing

        Dim wb2 As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(ssgx)
        Dim ws2 As Microsoft.Office.Interop.Excel.Worksheet = wb2.Worksheets(1)
        xlSourceRange = ws2.Range("A1:AN" & CStr(DataTable.Rows.Count))

        Dim rTotal As Int64 = DataTable.Select("").Length
        'r = 0
        'For Each row As DataRow In DataTable.Select("")
        '    r += 1
        '    ws.Range("O" & CStr(5 + r) & ":AX" & CStr(5 + r)).Value2 = row.ItemArray

        '    If r Mod 1000 = 0 Then
        '        ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
        '    End If
        'Next

        r = DataTable.Rows.Count

        xlDestRange = ws.Range("O" & CStr(5 + 1) & ":BB" & CStr(5 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        '' NOTE - THIS COPIES THE CLOSED INDICATOR FROM COL AL OF THE SSG XLS INTO COLUMN AZ OF THE XLS DATA TAB
        'xlSourceRange = ws2.Range("AM1:AM" & CStr(DataTable.Rows.Count))
        'xlDestRange = ws.Range("AZ" & CStr(5 + 1) & ":AZ" & CStr(5 + DataTable.Rows.Count))
        'xlSourceRange.Copy(xlDestRange)

        xlDestRange = Nothing
        xlSourceRange = Nothing
        ws2 = Nothing
        wb2 = Nothing
        ReleaseCOMObject(ws2)
        ReleaseCOMObject(wb2)

        ASCMAIN1.Progress("-", "Formulas")

        xlSourceRange = ws.Range("A6:N6")
        xlDestRange = ws.Range("A6:N" & CStr(5 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        ' COPY CHAIN
        xlSourceRange = ws.Range("BC6:BD6")
        xlDestRange = ws.Range("BC6:BD" & CStr(5 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        If ASCMAIN1.CLIENT = "INT" Then
            ws.Cells(3, 7).Value = Now
        Else
            ws.Cells(2, 1).Value = Now
        End If

        wb.Names.Add("PivotBase", "=DATA!$A$5:$BD$" & CStr(5 + DataTable.Rows.Count))



        Dim wx As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets.Add

        wx.Cells(1, 1).entirecolumn.numberformat = "@"

        wx.Cells(1, 1).VALUE = "TimeFrame"
        wx.Cells(1, 2).VALUE = "From"
        wx.Cells(1, 3).VALUE = "To"
        wx.Cells(1, 4).VALUE = "Weeks"

        Dim rwks As Integer = 1
        Weeks(rwks, wx, "TY_WK1", WKS(1, 0), WKS(1, 0))
        Weeks(rwks, wx, "TY_WK1", WKS(1, 0), WKS(1, 0))
        Weeks(rwks, wx, "LY_WK1", WKS(1, 1), WKS(1, 1))
        Weeks(rwks, wx, "TY_WK2", WKS(2, 0), WKS(2, 0))
        Weeks(rwks, wx, "LY_WK2", WKS(2, 1), WKS(2, 1))
        Weeks(rwks, wx, "TY_WK3", WKS(3, 0), WKS(3, 0))
        Weeks(rwks, wx, "LY_WK3", WKS(3, 1), WKS(3, 1))
        Weeks(rwks, wx, "TY_WK4", WKS(4, 0), WKS(4, 0))
        Weeks(rwks, wx, "LY_WK4", WKS(4, 1), WKS(4, 1))
        Weeks(rwks, wx, "TY_WK5", WKS(5, 0), WKS(5, 0))
        Weeks(rwks, wx, "LY_WK6", WKS(5, 1), WKS(5, 1))
        Weeks(rwks, wx, "TY_MTD", WKS(1, 0), WKS(REL_WEEK, 0))
        Weeks(rwks, wx, "LY_MTD", WKS(1, 1), WKS(REL_WEEK + WKS_ADJ, 1))
        Weeks(rwks, wx, "TY_HTD445", WKS(7, 0), WKS(REL_WEEK, 0))
        Weeks(rwks, wx, "LY_HTD445", WKS(7, 1), WKS(REL_WEEK + WKS_ADJ, 1))
        Weeks(rwks, wx, "TY_STD454", WKS(8, 0), WKS(REL_WEEK, 0))
        Weeks(rwks, wx, "LY_STD454", WKS(8, 1), WKS(REL_WEEK + WKS_ADJ, 1))
        Weeks(rwks, wx, "TY_YTD445", WKS(9, 0), WKS(REL_WEEK, 0))
        Weeks(rwks, wx, "LY_YTD445", WKS(9, 1), WKS(REL_WEEK + WKS_ADJ, 1))
        Weeks(rwks, wx, "TY_YTD445", WKS(10, 0), WKS(REL_WEEK, 0))
        Weeks(rwks, wx, "LY_YTD445", WKS(10, 1), WKS(REL_WEEK + WKS_ADJ, 1))

        wx.Name = "Week Ranges"
        wx.Visible = False


        'If chkUSE52WEEKS.Checked Then
        '    ws = wb.Worksheets("Event Doors")
        '    ws.Visible = False
        '    ws = wb.Worksheets("Top Doors")
        '    ws.Visible = False
        '    'ws = wb.Worksheets("Mth ByWeek 445")
        '    'ws.Visible = True
        'End If


        If ASCMAIN1.CLIENT = "INT" Then
            ' Do Nothing
        Else
            If ASCMAIN1.CLIENT = "SLP" Then
            Else


                ASCMAIN1.Progress("-", "Pivots")

                Dim Sheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim Pivot As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                For Each Sheet In wb.Worksheets
                    For Each Pivot In Sheet.PivotTables
                        Pivot.RefreshTable()
                        Pivot.Update()
                    Next
                Next

                If Pivot IsNot Nothing Then
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(Pivot)
                    Catch ex As Exception
                    End Try
                End If
                If Sheet IsNot Nothing Then
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(Sheet)
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Retail_Flash"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"

                Dim objOpt As Object = Nothing ' Missing.Value

                Dim FN As String = ASCMAIN1.Folders("Work") & XLS_FILENAME

                If My.Computer.FileSystem.FileExists(FN) Then
                    My.Computer.FileSystem.DeleteFile(FN)
                End If

                wb.SaveAs(FN, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then
                    MsgBox(ex.Message)
                End If
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

    Sub Weeks(ByRef rwks As Integer, WX As Microsoft.Office.Interop.Excel.Worksheet, COL As String, WK1 As String, WK2 As String)
        rwks += 1
        WX.Cells(rwks, 1).VALUE = COL
        WX.Cells(rwks, 2).VALUE = WK1
        WX.Cells(rwks, 3).VALUE = WK2
        If WK1 = WK2 Then
            WX.Cells(rwks, 4).VALUE = 1
        Else
            WX.Cells(rwks, 4).VALUE = ASCMAIN1.Week_Diff(WK1, WK2)
        End If
    End Sub
End Class