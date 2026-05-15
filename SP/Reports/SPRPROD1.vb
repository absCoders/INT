Public Class SPRPROD1

#Region "Declarations"
    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String

    Dim SEASON_YEAR_LY As String

    Dim SPTPROD1 As String
    Dim SPTCWRXW As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        If Format(Now, "yyyyMMdd") > "20170301" Then
            chkUseOldCB.Visible = False
        End If

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

        If chkUseOldCB.Checked Then
            ASCMAIN1.sql = "Select Distinct ICTCOLL1.COLLECTION_CODE, SPTCWRXC.CHECKBOOK from ICTCOLL1,SPTCWRXC" & vbCrLf _
                & " where SPTCWRXC.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
                & "   and (SPTCWRXC.COLLECTION_GENDER = 'U' or ICTCOLL1.COLLECTION_GENDER = SPTCWRXC.COLLECTION_GENDER)"
        Else
            ASCMAIN1.sql = "Select Distinct ICTCOLL1.COLLECTION_CODE, ICTCOLL0.CHECKBOOK from ICTCOLL1,ICTCOLL0" & vbCrLf _
                & " where ICTCOLL0.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
                & "   and ICTCOLL0.CHECKBOOK IS NOT NULL"
        End If

        SPTCWRXW = ASCMAIN1.Temp_Table

        ASCDATA1.ExecuteSQL("Alter Table " & SPTCWRXW & " Add Primary Key (COLLECTION_CODE)")

        Dim SQLW As String = ""
        SQLW &= SQLA_filter("CUST_CODE", "RSTRETL1")
        'SQLW &= SQLA_filter("SREP_CODE", "SOTORDR1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")

        ASCMAIN1.Progress("Now Creating Workbook")
        '"C:\Share\INT\Templates\" & Me.Name & ".xlsx" '
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
        Dim LYTW As String = ASCMAIN1.Week_Calc(RYW, -52)
        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")
        SEASON_YEAR = Mid(RYM, 1, 4)
        Dim MM As String = Mid(RYM, 5, 2)
        Dim M01 As String = IIf(Val(MM) > 6, "07", "01")
        SEASON_YEAR_LY = Format(Val(Mid(RYM, 1, 4) - 1), "0000")
        SEASON_CODE = SEASON_TYPE & SEASON_YEAR

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, CHECKBOOK, HC_CODE, COLLECTION_GENDER, RSSP_NAME" & vbCrLf _
            & ", Sum (EDI_SALES_TYTW) EDI_SALES_TYTW, Sum (EDI_SALES_LYTW) EDI_SALES_LYTW, Sum (EDI_SALES_TYTM) EDI_SALES_TYTM, Sum (EDI_SALES_LYTM) EDI_SALES_LYTM" & vbCrLf _
            & ", Sum (EDI_SALES_TYSTD) EDI_SALES_TYSTD, Sum (EDI_SALES_LYSTD) EDI_SALES_LYSTD, Sum (EDI_SALES_TYYTD) EDI_SALES_TYYTD, Sum (EDI_SALES_LYYTD) EDI_SALES_LYYTD" & vbCrLf _
            & ", Sum (RSC_SALES_TYTW) RSC_SALES_TYTW, Sum (RSC_HOURS_TYTW) RSC_HOURS_TYTW, Sum (RSC_SALES_TYTM) RSC_SALES_TYTM, Sum (RSC_HOURS_TYTM) RSC_HOURS_TYTM" & vbCrLf _
            & ", Sum (RSC_SALES_TYSTD) RSC_SALES_TYSTD, Sum (RSC_HOURS_TYSTD) RSC_HOURS_TYSTD, Sum (RSC_SALES_TYYTD) RSC_SALES_TYYTD, Sum (RSC_HOURS_TYYTD) RSC_HOURS_TYYTD" & vbCrLf _
            & ", Sum (RSC_SPEND_TYTW) RSC_SPEND_TYTW, Sum (RSC_SPEND_TYTM) RSC_SPEND_TYTM, Sum (RSC_SPEND_TYTS) RSC_SPEND_TYTS, Sum (RSC_SPEND_TYTY) RSC_SPEND_TYTY" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select X.*" & vbCrLf _
            & ", EDI.EDI_SALES_TYTW, EDI.EDI_SALES_LYTW, EDI.EDI_SALES_TYTM, EDI.EDI_SALES_LYTM" & vbCrLf _
            & ", EDI.EDI_SALES_TYSTD, EDI.EDI_SALES_LYSTD, EDI.EDI_SALES_TYYTD, EDI.EDI_SALES_LYYTD" & vbCrLf _
            & " from ICTBRAN1, ICTCOLL0, (" & vbCrLf _
            & "   Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK" & vbCrLf _
            & " , SUM (DECODE(OPS_YYYYWW,'" & RYW & "',AMT_SOLD,0)) EDI_SALES_TYTW" & vbCrLf _
            & " , SUM (DECODE(OPS_YYYYWW,'" & LYTW & "',AMT_SOLD,0)) EDI_SALES_LYTW" & vbCrLf _
            & " , SUM (DECODE(OPS_YYYYPP,'" & RYM & "',AMT_SOLD,0)) EDI_SALES_TYTM" & vbCrLf _
            & " , SUM (DECODE(OPS_YYYYPP,'" & SEASON_YEAR_LY & MM & "',AMT_SOLD,0)) EDI_SALES_LYTM" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & M01 & "'    and OPS_YYYYWW <= '" & RYW & "' THEN AMT_SOLD ELSE 0 END) EDI_SALES_TYSTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR_LY & M01 & "' and OPS_YYYYWW <= '" & LYTW & "' THEN AMT_SOLD ELSE 0 END)  EDI_SALES_LYSTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & "01" & "'    and OPS_YYYYWW <= '" & RYW & "' THEN AMT_SOLD ELSE 0 END) EDI_SALES_TYYTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR_LY & "01" & "' and OPS_YYYYWW <= '" & LYTW & "' THEN AMT_SOLD ELSE 0 END) EDI_SALES_LYYTD" & vbCrLf _
            & " , 0 RSC_SALES_TYTW, 0 RSC_HOURS_TYTW, 0 RSC_SALES_TYTM, 0 RSC_HOURS_TYTM, 0 RSC_SALES_TYSTD, 0 RSC_HOURS_TYSTD, 0 RSC_SALES_TYYTD, 0 RSC_HOURS_TYYTD" & vbCrLf _
            & " , 0 RSC_SPEND_TYTW, 0 RSC_SPEND_TYTM, 0 RSC_SPEND_TYTS, 0 RSC_SPEND_TYTY" & vbCrLf _
            & "  from RSTRETL1, ICTITEM1, ICTCOLL1, ARTCUST1, " & SPTCWRXW & " SPTCWRXW" & vbCrLf _
            & "  where RSTRETL1.OPS_YYYYPP between '" & SEASON_YEAR_LY & "01' and '" & RYM & "'" & vbCrLf _
            & "    and ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "    and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "    and ARTCUST1.CUST_CODE (+) = RSTRETL1.CUST_CODE" & vbCrLf _
            & "    and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "    and SPTCWRXW.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTRETL1") _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, SPTCWRXW.CHECKBOOK) EDI, (" & vbCrLf _
            & " Select X.CUST_CODE, X.CUST_STORE_NO, X.CHECKBOOK, ICTCOLL1.BRAND_CODE,  ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, X.RSSP_NAME" & vbCrLf _
            & "  , SUM (DECODE(OPS_YYYYWW,'" & RYW & "',SALES_AMT_TOTAL,0)) RSC_SALES_TYTW, SUM (DECODE(OPS_YYYYWW,'" & RYW & "',PAY_HOURS,0)) RSC_HOURS_TYTW" & vbCrLf _
            & "  , SUM (DECODE(OPS_YYYYPP,'" & RYM & "',SALES_AMT_TOTAL,0)) RSC_SALES_TYTM, SUM (DECODE(OPS_YYYYPP,'" & RYM & "',PAY_HOURS,0)) RSC_HOURS_TYTM " & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & M01 & "'    and OPS_YYYYWW <= '" & RYW & "' THEN SALES_AMT_TOTAL ELSE 0 END) RSC_SALES_TYSTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & M01 & "'    and OPS_YYYYWW <= '" & RYW & "' THEN PAY_HOURS ELSE 0 END) RSC_HOURS_TYSTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & "01" & "'    and OPS_YYYYWW <= '" & RYW & "' THEN SALES_AMT_TOTAL ELSE 0 END) RSC_SALES_TYYTD" & vbCrLf _
            & " , SUM (CASE WHEN OPS_YYYYPP >= '" & SEASON_YEAR & "01" & "'    and OPS_YYYYWW <= '" & RYW & "' THEN PAY_HOURS ELSE 0 END) RSC_HOURS_TYYTD" & vbCrLf _
            & ", Sum (DECODE(OPS_YYYYWW,'" & RYW & "',SPEND_AMT,0)) RSC_SPEND_TYTW" & vbCrLf _
            & ", Sum (DECODE(OPS_YYYYPP,'" & RYM & "',SPEND_AMT,0)) RSC_SPEND_TYTM" & vbCrLf _
            & ", Sum (CASE WHEN OPS_YYYYPP BETWEEN '" & Mid(RYM, 1, 4) & M01 & "' AND '" & RYM & "' THEN SPEND_AMT ELSE 0 END) RSC_SPEND_TYTS" & vbCrLf _
            & ", Sum (CASE WHEN OPS_YYYYPP BETWEEN '" & Mid(RYM, 1, 4) & "01" & "' AND '" & RYM & "' THEN SPEND_AMT ELSE 0 END) RSC_SPEND_TYTY" & vbCrLf _
            & " from ICTCOLL1, SPTCWRXC, ARTCUST1, (" & vbCrLf _
            & "  Select SPTPYXI1.CTRL_NO, SPTPYXI1.CTRL_LNO ,SPTPYXI1.CUST_CODE, SPTPYXI1.CUST_STORE_NO, SPTCWRXB.CHECKBOOK, SPTPYXI1.RSC_NAME RSSP_NAME" & vbCrLf _
            & ", 0 PAY_HOURS, 0 TOT_SALES_AMT, SPTPYXI1.OPS_YYYYPP, SPTPYXI1.OPS_YYYYWW, SPTPYXI2.SUB_BRAND, Sum(SPTPYXI2.SALES_AMT) SALES_AMT_TOTAL, 0 SPEND_AMT" & vbCrLf _
            & "  from SPTPYXI1, SPTPYXI2, SPTCWRXB" & vbCrLf _
            & "  where SPTPYXI1.CTRL_NO = SPTPYXI2.CTRL_NO" & vbCrLf _
            & "    and SPTPYXI1.CTRL_LNO = SPTPYXI2.CTRL_LNO" & vbCrLf _
            & "    and SPTCWRXB.SUB_BRAND = SPTPYXI2.SUB_BRAND" & vbCrLf _
            & "    and SPTPYXI1.OPS_YYYYPP between '" & SEASON_YEAR_LY & "01' and '" & RYM & "'" & vbCrLf _
            & "  group by SPTPYXI1.CTRL_NO, SPTPYXI1.CTRL_LNO, SPTPYXI1.CUST_CODE, SPTPYXI1.CUST_STORE_NO, SPTCWRXB.CHECKBOOK, SPTPYXI1.RSC_NAME" & vbCrLf _
            & ", SPTPYXI1.OPS_YYYYPP, SPTPYXI1.OPS_YYYYWW, SPTPYXI2.SUB_BRAND" & vbCrLf _
            & " UNION " & vbCrLf _
            & "  Select SPTPYXI1.CTRL_NO, SPTPYXI1.CTRL_LNO ,SPTPYXI1.CUST_CODE, SPTPYXI1.CUST_STORE_NO, SPTPYXI1.CHECKBOOK, SPTPYXI1.RSC_NAME RSSP_NAME" & vbCrLf _
            & ", SPTPYXI1.PAY_HOURS, SPTPYXI1.TOT_SALES_AMT, SPTPYXI1.OPS_YYYYPP, SPTPYXI1.OPS_YYYYWW, NULL SUB_BRAND, 0 SALES_AMT_TOTAL, SPTPYXI1.SPEND_AMT" & vbCrLf _
            & "  from SPTPYXI1" & vbCrLf _
            & "  Where OPS_YYYYPP between '" & SEASON_YEAR_LY & "01' and '" & RYM & "'" & vbCrLf _
            & ") X" & vbCrLf _
            & "  Where X.CHECKBOOK = SPTCWRXC.CHECKBOOK" & vbCrLf _
            & "  And  SPTCWRXC.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & "  And ARTCUST1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "X") _
            & "  Group by X.CUST_CODE, X.CUST_STORE_NO, X.CHECKBOOK, ICTCOLL1.BRAND_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.COLLECTION_GENDER, X.RSSP_NAME " & vbCrLf _
            & ") X " & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = X.BRAND_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE = X.HC_CODE" & vbCrLf _
            & "   and ICTCOLL0.HC_STATUS = 'A'" & vbCrLf _
            & "   and EDI.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & "   and EDI.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and EDI.CHECKBOOK (+) = X.CHECKBOOK" & vbCrLf _
            & ")  group by CUST_CODE, CUST_STORE_NO, BRAND_CODE, HC_CODE, CHECKBOOK, COLLECTION_GENDER, RSSP_NAME"

        SPTPROD1 = ASCMAIN1.Temp_Table

        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("Data")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.HC_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME" & vbCrLf _
            & ", SOTSREG1.REGION_DESC" & vbCrLf _
            & ", SOTSELL1.SELL_NAME" & vbCrLf _
            & ", SOTSELL1_AC.SELL_NAME AC" & vbCrLf _
            & ", X.RSSP_NAME" & vbCrLf _
            & ", X.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL0.HC_NAME" & vbCrLf _
            & ", X.CHECKBOOK" & vbCrLf _
            & ", EDI_SALES_TYTW, EDI_SALES_LYTW, RSC_SALES_TYTW, RSC_HOURS_TYTW" & vbCrLf _
            & ", EDI_SALES_TYTM, EDI_SALES_LYTM, RSC_SALES_TYTM, RSC_HOURS_TYTM" & vbCrLf _
            & ", EDI_SALES_TYSTD, EDI_SALES_LYSTD, RSC_SALES_TYSTD, RSC_HOURS_TYSTD" & vbCrLf _
            & ", EDI_SALES_TYYTD, EDI_SALES_LYYTD, RSC_SALES_TYYTD, RSC_HOURS_TYYTD" & vbCrLf _
            & ", RSC_SPEND_TYTW, RSC_SPEND_TYTM, RSC_SPEND_TYTS, RSC_SPEND_TYTY" & vbCrLf _
            & " from " & SPTPROD1 & " X" & vbCrLf _
            & ",ARTCUST1,ARTCUST2,SOTSELL1,SOTSREG1,ICTCOLL0,SOTSELL1 SOTSELL1_AC" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and ICTCOLL0.HC_CODE (+) = X.HC_CODE" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   and SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 4
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_NAME,HC_NAME,CUST_STORE_NO")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":AG" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=Data!$A$" & CStr(r0) & ":$AG$" & CStr(r0 + DataTable.Rows.Count))

        'ASCMAIN1.Progress("-", "Formulas")

        'xlSourceRange = ws.Range("AJ" & CStr(r0 + 1) & ":AJ" & CStr(r0 + 1))
        'xlDestRange = ws.Range("AJ" & CStr(r0 + 1) & ":AJ" & CStr(r0 + DataTable.Rows.Count))
        'xlSourceRange.Copy(xlDestRange)

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
                XLS_FILENAME = "Productivity_Report"
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