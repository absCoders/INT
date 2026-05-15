Public Class RSRVISW1

    Dim SEASON_CODE As String
    Dim RSTRBPL1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW0", ASCMAIN1.CYW, -52, 0, -1)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -104, 0, -52)
        Set_cmbYW("RYW2", ASCMAIN1.CYW, -104, 0, -52)
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
            If txtCUST_CODE.Text = "" Then
                EMsg &= vbCr & "You Must Specify a Customer"
            Else
                If LookUp("ARTCUST1", txtCUST_CODE.Text) Is Nothing Then
                    EMsg &= vbCr & "Customer Specified (" & txtCUST_CODE.Text & ") is Invalid"
                Else

                End If
            End If
            If Absx1.txtFor("HC_CODE").Text = "" Then
                EMsg &= vbCr & "You Must Specify a High Collection"
            End If
            If Absx1.cmbFor("RYW0").Value = "" Then
                EMsg &= vbCr & "You Must Specify a value for TY"
            End If
            If Absx1.cmbFor("RYW1").Value = "" Then
                EMsg &= vbCr & "You Must Specify a value for TY"
            End If
        End If
    End Sub

    Sub Create_Pivot()

        Dim Cust_Code As String = Absx1.txtFor("CUST_CODE").Text
        Dim HC_Code As String = Absx1.txtFor("HC_CODE").Text
        Dim TY As String = Mid(Replace(Absx1.cmbFor("RYW0").Value, "-", ""), 1, 6)
        Dim TYLW As String = ASCMAIN1.Week_Calc(TY, -1)
        Dim LY As String = Mid(Replace(Absx1.cmbFor("RYW1").Value, "-", ""), 1, 6)
        Dim LY_VISUAL As String = Mid(Replace(Absx1.cmbFor("RYW2").Value, "-", ""), 1, 6)
        Dim Summary_Row As Integer = 5
        Dim Earliest_LY As String = IIf(LY < LY_VISUAL, LY, LY_VISUAL)
        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx" '"C:\Share\INT\Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Data")
        ws = wb.Worksheets("Data")


        Dim Start_PP As String = IIf(Mid(TY, 5, 2) >= 27, "07", "01")
        Dim End_PP As String = IIf(Mid(TY, 5, 2) >= 27, "12", "06")

        Dim sqlRetailBudget As String = "Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO" & vbCrLf _
        & ", Sum(RSTBUDR1.BUDGET) as RETAIL_PLAN from RSTBUDR1, ICTCOLL1" & vbCrLf _
        & " Where RSTBUDR1.CUST_CODE = '" & Cust_Code & "'" & vbCrLf _
        & " And RSTBUDR1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCrLf _
        & " And ICTCOLL1.HC_CODE = '" & HC_Code & "'" & vbCrLf _
        & " And RSTBUDR1.OPS_YYYYPP between '" & Mid(TY, 1, 4) & Start_PP & "' and '" & Mid(TY, 1, 4) & End_PP & "'" & vbCrLf _
        & " Group by RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO"


        Dim sqlRetailSales As String = "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
        & " , SUM (CASE WHEN OPS_YYYYWW BETWEEN '" & Mid(TY, 1, 4) & IIf(Mid(TY, 5, 2) >= 27, "27", "01") & "' AND '" & TY & "' THEN AMT_SOLD ELSE 0 END) STD" & vbCrLf _
        & " , SUM (CASE WHEN OPS_YYYYWW BETWEEN '" & Mid(TY, 1, 4) & "01' AND '" & TY & "' THEN AMT_SOLD ELSE 0 END) YTD" & vbCrLf _
        & " , SUM (DECODE(OPS_YYYYWW,'" & TY & "',AMT_SOLD,0)) TYLW" & vbCrLf _
        & " , SUM (DECODE(OPS_YYYYWW,'" & LY & "',AMT_SOLD,0)) LYW" & vbCrLf _
        & " , SUM (DECODE(OPS_YYYYWW,'" & LY_VISUAL & "',AMT_SOLD,0)) LYLW" & vbCrLf _
        & " From RSTRETL1, ICTITEM1, ICTCOLL1 Where RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
        & " AND RSTRETL1.CUST_CODE = '" & Cust_Code & "'" & vbCrLf _
        & " AND RSTRETL1.OPS_YYYYWW >='" & Earliest_LY & "' AND RSTRETL1.OPS_YYYYWW <='" & TY & "'" & vbCrLf _
        & " AND ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCrLf _
        & " And ICTCOLL1.HC_CODE = '" & HC_Code & "'" & vbCrLf _
        & " Group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO"

        ASCMAIN1.sql = " Select ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & ", CASE WHEN SATAUTH1.OPS_YYYYPP_OPENED > '" & ASCMAIN1.CYP & "' THEN SATAUTH1.OPS_YYYYPP_OPENED ELSE NULL END FUTURE_CLOSED" & vbCrLf _
            & ", CASE WHEN SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "' THEN SATAUTH1.OPS_YYYYPP_CLOSED ELSE NULL END FUTURE_CLOSED" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_STATE, SOTSELL1.SELL_NAME" & vbCrLf _
            & ",  Decode(SOTSREG1.REGION_DESC,Null,Nvl(ARTCUST2.SELL_CODE,'No Code'),REGION_DESC)" & vbCrLf _
            & ", Nvl(Y.RETAIL_PLAN,0), 0 as PERC_TO_TOTAL_COL_G, X.STD, 0 as PERC_TO_TOTAL_COL_I, X.YTD, 0 as PERC_TO_TOTAL_COL_K, X.TYLW as TYLW,  0 as PERC_TO_TOTAL_COL_M " & vbCrLf _
            & ", LYW as LYLW_SALES, 0 as PERC_TO_TOTAL_COL_O, LYLW as LY_VISUAL,  0 as PERC_TO_TOTAL_COL_Q" & vbCrLf _
            & " from ARTCUST2, SATAUTH1, SOTSELL1, SOTSDSC1, SOTSREG1" & vbCrLf _
            & ", (" & sqlRetailSales & " ) X" & vbCrLf _
            & ", (" & sqlRetailBudget & ") Y " & vbCrLf _
            & " where SATAUTH1.CUST_CODE = '" & Cust_Code & "'" & vbCrLf _
            & "   and SATAUTH1.HC_CODE = '" & HC_Code & "'" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED is Not Null" & vbCrLf _
            & "   and (SATAUTH1.OPS_YYYYPP_CLOSED is Null or SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _
            & "   and X.CUST_CODE (+) = SATAUTH1.CUST_CODE" & vbCrLf _
            & "   and X.CUST_STORE_NO (+) = SATAUTH1.CUST_STORE_NO" & vbCrLf _
            & "   and Y.CUST_CODE (+) = SATAUTH1.CUST_CODE" & vbCrLf _
            & "   and Y.CUST_STORE_NO (+) = SATAUTH1.CUST_STORE_NO" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = SATAUTH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SATAUTH1.CUST_STORE_NO" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
            & "   and SOTSDSC1.SDS_CODE (+) = ARTCUST2.SDS_CODE"
        DataTable = ASCDATA1.GetDataTable

        r0 = 6
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "CUST_STORE_NAME")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":S" & CStr(r0 + r)).Value2 = row.ItemArray


            If r Mod 1000 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next

        Dim CH = 7 ' number of header columns before the 1st Total column (Retail Plan)
        For Each c As Integer In New Integer() {1, 3, 5, 7, 9, 11, 13, 21, 22, 23} ' Total Columns
            Dim XC As String = Excel_Cell(-1, CH + c)
            ws.Cells(Summary_Row, CH + c).Value = "=SUM(" & XC & r0 + 1 & ":" & XC & r0 + DataTable.Rows.Count & ")"
        Next

        ASCMAIN1.Progress("-", "Formulas")
        Dim cc2 As Integer
        Dim Col1 As String = ""
        Dim Col2 As String = ""

        For Each c As Integer In New Integer() {1, 3, 5, 7, 9, 11} ' Total Columns with % to Total
            Dim Column_Header As String = Excel_Cell(-1, CH + c)
            ws.Range(Column_Header & r0 + 1 & ":" & Column_Header & r0 + DataTable.Rows.Count).NumberFormat = "$#,##0.00"

            Col1 = "$" & Column_Header & "$" & Summary_Row
            Col2 = Column_Header & r0 + 1

            cc2 = CH + c + 1

            ws.Cells(r0 + 1, cc2).Value = "=(" & Col2 & "/" & Col1 & ")"
            xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc2))
            xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc2), Excel_Cell(r0 + DataTable.Rows.Count, cc2))
            xlSourceRange.Copy(xlDestRange)

            Dim Column_Header_pct As String = Excel_Cell(-1, cc2)
            ws.Range(Column_Header_pct & r0 + 1 & ":" & Column_Header_pct & r0 + DataTable.Rows.Count).NumberFormat = "###0.00%"
        Next

        For i As Integer = CH + 13 To CH + 23

            cc2 = i
            xlSourceRange = ws.Range(Excel_Cell(r0 + 1, cc2))
            xlDestRange = ws.Range(Excel_Cell(r0 + 1, cc2), Excel_Cell(r0 + DataTable.Rows.Count, cc2))
            xlSourceRange.Copy(xlDestRange)
        Next

        ws.Cells(1, 1).Value = Now
        ws.Cells(4, 1).Value = "High Collection: " & HC_Code
        ws.Cells(4, 1).Font.Bold = True

        xlSourceRange = ws.Range("A4", "D4")
        xlSourceRange.Merge()

        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Visual_Week"
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

    Private Sub SplitContainer5_Panel1_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer5.Panel1.Paint

    End Sub
End Class