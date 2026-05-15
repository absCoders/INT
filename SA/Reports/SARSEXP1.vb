Public Class SARSEXP1

#Region "Declarations"
    Dim SATSEXP1 As String
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
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
        SQLW &= SQLA_filter("CUST_CODE", "SOTINVH2")
        SQLW &= SQLA_filter("SREP_CODE", "ARTCUST1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")

        ASCMAIN1.Progress("Now Creating Workbook")

        'Dim FILENAME As String = "S:\AHA\Templates\" & Me.Name & ".xlsx"
        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "\Templates\" & Me.Name & ".xlsx"

        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)


        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)

        Dim YPs(1, 12) As String
        For i As Integer = 1 To 12
            YPs(0, i) = Mid(RYP, 1, 4) & Format(i, "00")
            YPs(1, i) = Format(Val(Mid(RYP, 1, 4)) - 1, "0000") & Format(i, "00")
        Next

        Dim sqlsum As String = ""
        Dim sqlCs As String = ""
        Dim sqlSOTINVH2 As String = ""
        Dim sqlCOST As String = "SOTINVH2.ORDR_QTY_SHIP * DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST)"

        For Each Y As String In New String() {"LY", "TY"}
            Dim j As Integer = IIf(Y = "TY", 0, 1)
            Dim sqlTOT As String = ""
            Dim sqlYTD As String = ""
            For I As Integer = 1 To 12
                Dim C As String = Y & "_M" & Format(I, "00")
                sqlCs &= ", " & C
                sqlsum &= ", Sum (" & C & ") " & C
                If I <= Val(Mid(RYP, 5, 2)) Then
                    sqlYTD &= "+NVL(" & C & ",0)"
                End If
                sqlTOT &= "+NVL(" & C & ",0)"

                sqlSOTINVH2 &= ", Sum (Decode(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YPs(j, I) & "'," & sqlCOST & ",0)) " & C & vbCrLf
            Next
            sqlsum &= vbCrLf
            sqlCs &= ", " & Y & "_" & "YTD"
            sqlsum &= ", Sum (" & Mid(sqlYTD, 2) & ") " & Y & "_" & "YTD" & vbCrLf
            sqlCs &= ", " & Y & "_" & "TOT"
            sqlsum &= ", Sum (" & Mid(sqlTOT, 2) & ") " & Y & "_" & "TOT" & vbCrLf
        Next

        ASCMAIN1.Progress("Work Table")
        ASCMAIN1.sql = "Select DATA_TYPE, CUST_CODE, COLLECTION_CODE" & vbCrLf _
            & sqlsum _
            & " from (" & vbCrLf _
            & "Select 'RTLSLS' DATA_TYPE, RSTRETL1.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & Replace(Replace(Replace(sqlSOTINVH2, sqlCOST, "AMT_SOLD"), "SOTINVH2", "RSTRETL1"), "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP") _
            & " from RSTRETL1, ICTITEM1, ARTCUST1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = RSTRETL1.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "SOTINVH2", "RSTRETL1") _
            & " group by RSTRETL1.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'RTLBUD' DATA_TYPE, RSTBUDR1.CUST_CODE, RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & Replace(Replace(Replace(sqlSOTINVH2, sqlCOST, "BUDGET"), "SOTINVH2", "RSTBUDR1"), "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP") _
            & " from RSTBUDR1, ICTCOLL1, ARTCUST1" & vbCrLf _
            & " where RSTBUDR1.OPS_YYYYPP between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = RSTBUDR1.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "SOTINVH2", "RSTBUDR1") _
            & " group by RSTBUDR1.CUST_CODE, RSTBUDR1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select DECODE(ICTITEM1.PROD_CODE,'TESTER','TSTCGS','SAMCGS') DATA_TYPE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & sqlSOTINVH2 _
            & " from SOTINVH2, ICTITEM1, ARTCUST1, SOTINVH1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and ICTITEM1.PROD_CODE IN ('TESTER','NCDELUXE','NCS')" & vbCrLf _
            & "   and SOTINVH1.EVENT_CODE IS NULL" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & SQLW _
            & " group by DECODE(ICTITEM1.PROD_CODE,'TESTER','TSTCGS','SAMCGS'), SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'EVTCGS' || SOTINVH1.EVENT_CODE DATA_TYPE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & sqlSOTINVH2 _
            & " from SOTINVH2, ICTITEM1, ARTCUST1, SOTINVH1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH1.EVENT_CODE IS NOT NULL" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & SQLW _
            & " group by 'EVTCGS' || SOTINVH1.EVENT_CODE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select DECODE(ICTITEM1.PROD_CODE,'TESTER','TSTSLS','SAMSLS') DATA_TYPE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & Replace(sqlSOTINVH2, "ITEM_UNIT_COST", "ORDR_UNIT_PRICE") _
            & " from SOTINVH2, ICTITEM1, ARTCUST1, SOTINVH1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE <> 0" & vbCrLf _
            & "   and ICTITEM1.PROD_CODE IN ('TESTER','NCDELUXE','NCS')" & vbCrLf _
            & "   and SOTINVH1.EVENT_CODE IS NULL" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & SQLW _
            & " group by DECODE(ICTITEM1.PROD_CODE,'TESTER','TSTSLS','SAMSLS'), SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'EVTSLS' || SOTINVH1.EVENT_CODE DATA_TYPE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & Replace(sqlSOTINVH2, "DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST)", "SOTINVH2.ORDR_UNIT_PRICE") _
            & " from SOTINVH2, ICTITEM1, ARTCUST1, SOTINVH1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_PRICE <> 0" & vbCrLf _
            & "   and SOTINVH1.EVENT_CODE IS NOT NULL" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & SQLW _
            & " group by 'EVTSLS' || SOTINVH1.EVENT_CODE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'RTNSLS' DATA_TYPE, SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & Replace(sqlSOTINVH2, "DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST)", "SOTINVH2.ORDR_UNIT_PRICE") _
            & " from SOTINVH2, ICTITEM1, ARTCUST1, SOTINVH1" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'C'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & SQLW _
            & " group by SOTINVH2.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'COMADV' DATA_TYPE, SPTACOMB.CUST_CODE, SPTACOMB.COLLECTION_CODE" & vbCrLf _
            & Replace(Replace(Replace(sqlSOTINVH2, sqlCOST, "SPTACOMB.AMT_COMM"), "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP"), "SOTINVH2", "SPTACOMB") _
            & " from SPTACOMB, ARTCUST1" & vbCrLf _
            & " where SPTACOMB.OPS_YYYYPP between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "SOTINVH2", "SPTACOMB") _
            & " group by SPTACOMB.CUST_CODE, SPTACOMB.COLLECTION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'PROEXP' DATA_TYPE, SPTCOOP1.CUST_CODE, SPTCOOP3.COLLECTION_CODE" & vbCrLf _
            & Replace(Replace(Replace(sqlSOTINVH2, sqlCOST, "SPTCOOP3.DIST_AMT"), "ORDR_YYYYPP_UPDATED", "OPS_YYYYPP"), "SOTINVH2", "SPTCOOP1") _
            & " from SPTCOOP3, ARTCUST1, SPTCOOP1" & vbCrLf _
            & " where SPTCOOP1.OPS_YYYYPP between '" & YPs(1, 1) & "' and '" & YPs(0, 12) & "'" & vbCrLf _
            & "   and SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
            & Replace(SQLW, "SOTINVH2", "SPTCOOP1") _
            & " group by SPTCOOP1.CUST_CODE, SPTCOOP3.COLLECTION_CODE" & vbCrLf _
            & ")  group by DATA_TYPE, CUST_CODE, COLLECTION_CODE"
        SATSEXP1 = ASCMAIN1.Temp_Table

        ' PROBABLY NEED TO PRORATE SPTACOMC BY COLLECTION IN SPTACOMB INSTEAD OF USING SPTACOMB, IN ABOVE SQL

        Dim rTotal As Int64
        Dim r0 As Integer

        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("MainData")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.DATA_TYPE" & vbCrLf _
            & ", X.CUST_CODE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", X.COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", SOTTCLS1.TRADE_CLASS_DESC" & vbCrLf _
            & ", ARTCUST1.SREP_CODE" & vbCrLf _
            & ", SOTSREP1.SREP_NAME" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & sqlCs & vbCrLf _
            & " from " & SATSEXP1 & " X" & vbCrLf _
            & ",ICTBRAN1,ARTCUST1,SOTTCLS1,SOTSREP1,ICTCOLL1" & vbCrLf _
            & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE = ARTCUST1.SREP_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = X.COLLECTION_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 4
        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "DATA_TYPE,CUST_CODE,COLLECTION_CODE")
            r += 1
            ws.Range("A" & CStr(r0 + r) & ":AM" & CStr(r0 + r)).Value2 = row.ItemArray
            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=MainData!$A$" & CStr(r0) & ":$AM$" & CStr(r0 + DataTable.Rows.Count))

        ASCMAIN1.Progress("-", "Formulas")

        'xlSourceRange = ws.Range("AI" & CStr(r0 + 1) & ":AI" & CStr(r0 + 1))
        'xlDestRange = ws.Range("AI" & CStr(r0 + 1) & ":AI" & CStr(r0 + DataTable.Rows.Count))
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
                XLS_FILENAME = "Sales_Expense_Summary"
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