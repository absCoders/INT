Imports Microsoft.Office.Interop
Public Class SARMMAL1

#Region "General Declarations"

    Dim XL As Microsoft.Office.Interop.Excel.Application = Nothing
    Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Dim range As Microsoft.Office.Interop.Excel.Range = Nothing
    Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
    Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

    Dim SATMMAL1 As String

    Structure YP_parms
        Dim FYP As String
        Dim LYP As String
        Dim FYPLEGEND As String
        Dim LYPLEGEND As String
        Dim RANGE As String
    End Structure

    Dim YP As New Dictionary(Of String, YP_parms)

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        RWU = "U"
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)

        ' Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        ' Get Run-Time options

        Dim CAL As String = Absx1.optFor("OPTCAL").Value

        ' Prepare Work File with Data from Server

        ASCMAIN1.Progress("Gathering Report Data", "")

        MyBase.Get_SQL("*")


        Dim iDATA_TYPE As Integer = 0
        For Each DATA_TYPE As String In New String() {"YTD_LY", "YTD_TY", "FYPY", "FYLY"}
            iDATA_TYPE += 1

            ' Set up First and Last YPs for TY TYD

            Dim FYP As String = ""
            If CAL = "O" Then
                FYP = Mid(RYP, 1, 4) & "01"
            Else
                If Mid(RYP, 1, 4) = "01" Then
                    FYP = Format(Val(Mid(RYP, 1, 4)) - 1, "0000") & "02"
                Else
                    FYP = Mid(RYP, 1, 4) & "02"
                End If
            End If

            Dim LYP As String = RYP

            Dim FYYYY As Integer = Val(Mid(FYP, 1, 4))
            Dim FPP As String = Mid(FYP, 5, 2)
            Dim LYYYY As Integer = Val(Mid(LYP, 1, 4))
            Dim LPP As String = Mid(LYP, 5, 2)

            Select Case DATA_TYPE
                Case "YTD_LY"
                    FYP = Format(FYYYY - 1, "0000") & FPP
                    LYP = Format(LYYYY - 1, "0000") & LPP
                Case "YTD_TY"
                    ' DO NOTHING
                Case "FYPY"
                    FYP = Format(FYYYY - 2, "0000") & FPP
                    LYP = ASCMAIN1.Period_Calc(FYP, 11)
                Case "FYLY"
                    FYP = Format(FYYYY - 1, "0000") & FPP
                    LYP = ASCMAIN1.Period_Calc(FYP, 11)
            End Select

            Dim sql_filter As String = " and RSTRETL2.OPS_YYYYPP between '" & FYP & "' AND '" & LYP & "'"

            ASCMAIN1.sql = "" _
                & "Select '" & DATA_TYPE & "' DATA_TYPE" & vbCrLf _
                & ", RSTRETL2.CUST_CODE, RSTRETL2.CUST_STORE_NO" & vbCrLf _
                & ", ICTBRAN1.SALES_DIVISION_CODE, RSTRETL2.COLLECTION_CODE" & vbCrLf _
                & ", Sum (RSTRETL2.RETAIL_SALES) SLS" _
                & " from RSTRETL2" & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by RSTRETL2.CUST_CODE, RSTRETL2.CUST_STORE_NO, " _
                & "ICTBRAN1.SALES_DIVISION_CODE, RSTRETL2.COLLECTION_CODE"

            If iDATA_TYPE = 1 Then
                SATMMAL1 = ASCMAIN1.Temp_Table
            Else
                ASCMAIN1.sql = "Insert into " & SATMMAL1 & " " & ASCMAIN1.sql
                ASCDATA1.ExecuteSQL()
            End If

            Dim YPP As New YP_parms
            YPP.FYP = FYP
            YPP.LYP = LYP
            YPP.FYPLEGEND = LookUp("GLTPARM2", FYP).Item("LEGEND")
            YPP.LYPLEGEND = LookUp("GLTPARM2", LYP).Item("LEGEND")
            YPP.RANGE = Mid(YPP.FYPLEGEND, 10, 6) & "-" & Mid(YPP.LYPLEGEND, 10, 6)

            YP.Add(DATA_TYPE, YPP)
        Next


        ' Output to Excel

        ASCMAIN1.Progress("Excel", "")
        Load_Excel()

    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        ' Output to Excel
        ASCMAIN1.Progress("Excel", "")
        Load_Excel()

    End Sub

    Public Overrides Sub Print_Report()
        'SUBT = ""
        'CR_params.Add("SUBT", SUBT)
        'Generate_Report(RPT, , SUBT)

        ' KICK OUT THE EXCEL SPREADSHEET

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        'ASCMAIN1.sql = "Insert into SPTDCOMB Select * from "
        'ASCDATA1.ExecuteSQL(sql)

    End Sub

    Sub Load_Excel()

        ' Create Workbook

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"

        XL = New Microsoft.Office.Interop.Excel.Application
        wb = XL.Workbooks.Open(FILENAME)


        ' Get COLLECTIONS, in BRAND order, for Pivot Table columns

        Dim CODES As String = ""
        Dim COLLECTION_CODEs As New List(Of String)
        Dim BRAND_CODEs As New List(Of String)
        Dim COLLECTION_BRANDs As New List(Of String)
        Dim COLLECTION_GENDERs As New List(Of String)

        ASCMAIN1.sql = "" _
            & "Select COLLECTION_CODE, COLLECTION_NAME, COLLECTION_GENDER, BRAND_CODE" _
            & " from ICTCOLL1 where COLLECTION_CODE" _
            & " in (Select Distinct COLLECTION_CODE from " & SATMMAL1 & " SATMMAL1)"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select _
            ("", "BRAND_CODE,COLLECTION_GENDER,COLLECTION_CODE")

            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
            Dim BRAND_CODE As String = row.Item("BRAND_CODE")
            Dim COLLECTION_GENDER As String = row.Item("COLLECTION_GENDER")

            CODES &= ",'" & COLLECTION_CODE & "'"
            COLLECTION_CODEs.Add(COLLECTION_CODE)
            If Not BRAND_CODEs.Contains(BRAND_CODE) Then BRAND_CODEs.Add(BRAND_CODE)
            COLLECTION_BRANDs.Add(BRAND_CODE)
            COLLECTION_GENDERs.Add(COLLECTION_GENDER)
        Next



        Dim maxb As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 2 + 1)
        Dim maxA As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 2 + 1 + 1)



        ' MainData

        ASCMAIN1.Progress("MainData", "")
        ws = wb.Sheets("MainData")

        ASCMAIN1.sql = "" _
            & "Select ARTCUST2.CUST_CODE || '-' || ARTCUST2.CUST_STORE_NO DOORKEY" & vbCrLf _
            & ", X.SALES_DIVISION_CODE COMPANY, ARTCUST2.DMA_CODE DMA_CODE, ASTZDDMA.NAME DMA_NAME" & vbCrLf _
            & ", ARTMALL1.MALL_NAME, ARTCUST2.CUST_CODE CHAIN, ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
            & ", ARTCUST2.SELL_CODE CFG_AE, SOTSELL1.SELL_NAME CFG_AENAME, SOTSREG1.REGION_DESC CFG_ASDNAME" & vbCrLf _
            & ", X.FYLY, X.FYPY, X.YTD_TY, X.YTD_LY" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME || '(' ||" _
            & "INITCAP(SUBSTR(SELL_NAME,1,INSTR(SELL_NAME,' '))) || SUBSTR(SELL_NAME,INSTR(SELL_NAME,' ')+1,1)" & vbCrLf _
            & "|| ')' DOORNAME2" & vbCrLf _
            & ", 0 LY_COUNT, 0 TY_COUNT" & vbCrLf _
            & ", INITCAP(SUBSTR(SELL_NAME,1,INSTR(SELL_NAME,' '))) || SUBSTR(SELL_NAME,INSTR(SELL_NAME,' ')+1,1) ABBR" & vbCrLf _
            & " from ARTCUST2,ARTMALL1,SOTSELL1,SOTSREG1,ASTZDDMA" _
            & ", (Select SALES_DIVISION_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
            & ", Sum (Decode(DATA_TYPE,'FYLY',SLS,0)) FYLY" & vbCrLf _
            & ", Sum (Decode(DATA_TYPE,'FYPY',SLS,0)) FYPY" & vbCrLf _
            & ", Sum (Decode(DATA_TYPE,'YTD_TY',SLS,0)) YTD_TY" & vbCrLf _
            & ", Sum (Decode(DATA_TYPE,'YTD_LY',SLS,0)) YTD_LY" & vbCrLf _
            & " from " & SATMMAL1 & " SATMMAL1 group by SALES_DIVISION_CODE, CUST_CODE, CUST_STORE_NO) X" _
            & " where ARTMALL1.MALL_CODE (+) = ARTCUST2.MALL_CODE" _
            & "   and ASTZDDMA.DMA (+) = ARTCUST2.DMA_CODE" _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" _
            & "   and ARTCUST2.CUST_CODE (+) = X.CUST_CODE" _
            & "   and ARTCUST2.CUST_STORE_NO (+) = X.CUST_STORE_NO"
        Dim tbl_MainData As DataTable = Load_DataTable(7, 1, "DOORKEY,COMPANY")

        If tbl_MainData.Rows.Count = 0 Then
            MsgBox("No Data")
            Exit Sub

        End If

        ws.Range("K7").Formula = "=IFERROR(VLOOKUP($A7,INDIRECT(K$2),K$3,FALSE),0)"
        ws.Range("L7").Formula = "=IFERROR(VLOOKUP($A7,INDIRECT(L$2),L$3,FALSE),0)"
        ws.Range("M7").Formula = "=IFERROR(VLOOKUP($A7,INDIRECT(M$2),M$3,FALSE),0)"
        ws.Range("N7").Formula = "=IFERROR(VLOOKUP($A7,INDIRECT(N$2),N$3,FALSE),0)"

        xlSourceRange = ws.Range("K7:N7")
        xlDestRange = ws.Range("K7:N" & CStr(6 + tbl_MainData.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        ws.Range("P7").Formula = "=IF(N7>0,1,0)"
        ws.Range("Q7").Formula = "=IF(M7>0,1,0)"

        xlSourceRange = ws.Range("P7:Q7")
        xlDestRange = ws.Range("P7:Q" & CStr(6 + tbl_MainData.Rows.Count))
        xlSourceRange.Copy(xlDestRange)

        For i As Integer = 1 To 18
            Dim LIST_NAME As String = ws.Range(Excel_Cell(6, i)).Value
            Dim C As String = Excel_Cell(0, i)
            wb.Names.Add(LIST_NAME, "=MainData!$" & C & "$6:$" & C & "$" & CStr(6 + tbl_MainData.Rows.Count))
        Next

        wb.Names.Add("Pivotbase", "=MainData!$A$6:$Q$" & CStr(6 + tbl_MainData.Rows.Count))


        'ws.Cells("A4:Q4").Copy(worksheet.Cells("A4:Q" & CStr(3 + DataTable.Rows.Count)))
        ws.Range("D1").Value = Now


        ' Data Tabs

        ASCMAIN1.Progress("Data Tabs", "")

        For Each DATA_TYPE As String In New String() {"YTD_LY", "YTD_TY", "FYPY", "FYLY"}

            ASCMAIN1.Progress("-", DATA_TYPE)
            ws = wb.Sheets("Data_" & DATA_TYPE)

            ws.Range("B2").Value = YP(DATA_TYPE).RANGE
            ws.Range("B3").Value = DATA_TYPE

            If BRAND_CODEs.Count > 1 Then
                For B As Integer = 2 To BRAND_CODEs.Count
                    ws.Range("F1").EntireColumn.Insert( _
                        Excel.XlInsertShiftDirection.xlShiftToRight,
                        Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                Next
            End If

            If COLLECTION_CODEs.Count > 1 Then
                For B As Integer = 2 To COLLECTION_CODEs.Count
                    ws.Range("E1").EntireColumn.Insert( _
                        Excel.XlInsertShiftDirection.xlShiftToRight,
                        Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                Next
            End If

            ws.Range(Excel_Cell(3, 4)).Resize(1, COLLECTION_GENDERs.Count).Value = COLLECTION_GENDERs.ToArray
            ws.Range(Excel_Cell(4, 4)).Resize(1, COLLECTION_BRANDs.Count).Value = COLLECTION_BRANDs.ToArray
            ws.Range(Excel_Cell(7, 4)).Resize(1, COLLECTION_CODEs.Count).Value = COLLECTION_CODEs.ToArray

            ws.Range(Excel_Cell(4, 3 + COLLECTION_CODEs.Count + 1)).Resize(1, BRAND_CODEs.Count).Value = BRAND_CODEs.ToArray
            ws.Range(Excel_Cell(7, 3 + COLLECTION_CODEs.Count + 1)).Resize(1, BRAND_CODEs.Count).Value = BRAND_CODEs.ToArray

            ASCMAIN1.sql = "Select * from (" _
                & "Select SATMMAL1.CUST_CODE || '-' || SATMMAL1.CUST_STORE_NO DOORKEY" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
                & ", SATMMAL1.SALES_DIVISION_CODE COMPANY, SATMMAL1.COLLECTION_CODE, SLS" & vbCrLf _
                & " from " & SATMMAL1 & " SATMMAL1, ARTCUST2" & vbCrLf _
                & " where SATMMAL1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SATMMAL1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO (+) = SATMMAL1.CUST_STORE_NO" & vbCrLf _
                & ") Pivot " & vbCrLf _
                & "(" & vbCrLf _
                & "  Sum(SLS)" & vbCrLf _
                & "  for COLLECTION_CODE" & vbCrLf _
                & "  in (" & Mid(CODES, 2) & ")" & vbCrLf _
                & ")"

            Dim tbl As DataTable = Load_DataTable(8, 1, "DOORKEY,COMPANY")

            Dim maxc As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count)
            For b As Integer = 1 To BRAND_CODEs.Count
                Dim maxc1 As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + b)
                ws.Range(Excel_Cell(8, 3 + COLLECTION_CODEs.Count + b)).Formula = "=SUMIF($D$4:$" & maxc & "$4," & maxc1 & "$4,$D8:$" & maxc & "8)"
            Next

            Dim maxcM As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 1)
            ws.Range(maxcM & "8").Formula = "=SUMIF($D$3:$" & maxc & "$3," & maxcM & "$3,$D8:$" & maxc & "8)"
            Dim maxcW As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 2)
            ws.Range(maxcW & "8").Formula = "=SUMIF($D$3:$" & maxc & "$3," & maxcW & "$3,$D8:$" & maxc & "8)"
            Dim maxcT As String = Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 3)
            ws.Range(maxcT & "8").Formula = "=SUM($D8:$" & maxc & "8)"

            xlSourceRange = ws.Range( _
                Excel_Cell(8, 3 + COLLECTION_CODEs.Count + 1), _
                Excel_Cell(8, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 2 + 1 + 1))
            xlDestRange = ws.Range( _
                Excel_Cell(8, 3 + COLLECTION_CODEs.Count + 1), _
                Excel_Cell(7 + tbl.Rows.Count, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 2 + 1 + 1))
            xlSourceRange.Copy(xlDestRange)

            ws.Range("D2:D2").Formula = "=SUM(D8:D" & CStr(7 + tbl.Rows.Count) & ")/1000"
            xlSourceRange = ws.Range("D2:D2")
            xlDestRange = ws.Range("D2:" & maxb & "2")
            xlSourceRange.Copy(xlDestRange)

            wb.Names.Add("Header_" & DATA_TYPE, "=Data_" & DATA_TYPE & "!$A$7:$" & maxb & "$7")
            wb.Names.Add("Data_" & DATA_TYPE, "=Data_" & DATA_TYPE & "!$A$3:$" & maxA & "$" & CStr(7 + tbl.Rows.Count))
        Next


        ' Top 30 Markets

        ASCMAIN1.Progress("Top 30 Markets", "")

        ws = wb.Worksheets("Top 30 Markets")
        With ws.Range("B1").Validation
            .Delete()
            .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList, _
                 AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop, _
                 Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween, _
                 Formula1:="=Data_YTD_TY!$D$7:$" & Excel_Cell(0, 3 + COLLECTION_CODEs.Count + BRAND_CODEs.Count + 3) & "$7")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select a Brand or Collection"
            .ErrorTitle = ""
            .InputMessage = "Hi Mom"
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With

        ws.Range("O2").Value = YP("YTD_TY").RANGE
        Dim CAL As String = Absx1.optFor("OPTCAL").Value
        ws.Range("O3").Value = Absx1.optFor("OPTCAL").Text & " (" & IIf(CAL = "R", "454", "445") & ") Calendar"
        ws.Range("B3").Value = txtDescription.Text


        ASCMAIN1.Progress("Refreshing Pivot Tables", "")
        ' this actually may work - wjz disabled all calls to excel.run("resetdata") - but this coding technique may be better - need to test
        'For Each ws2 As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
        '    For Each T As Microsoft.Office.Interop.Excel.QueryTable In ws2.QueryTables
        '        T.BackgroundQuery = False
        '    Next
        'Next

        'wb.RefreshAll()

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Market Mall Analysis"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"
                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
                wb.Close(False, objOpt, objOpt)
                success = True

            Catch ex As Exception
                Stop
            End Try
        Loop

        XL.Quit()
        ws = Nothing
        wb = Nothing
        XL = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(XL)

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Function Load_DataTable(Rx As Integer, Cx As Integer, Optional OrderBy As String = "") As DataTable

        Dim tbl As DataTable = ASCDATA1.GetDataTable

        Dim DataArray(,) As Object
        ReDim DataArray(tbl.Rows.Count, tbl.Columns.Count)
        Dim rows() As DataRow = tbl.Select("", OrderBy)
        For r As Integer = 0 To tbl.Rows.Count - 1
            For DC As Integer = 0 To tbl.Columns.Count - 1
                DataArray(r, DC) = rows(r).Item(DC)
            Next
        Next

        If tbl.Rows.Count <> 0 Then
            ws.Range(Excel_Cell(Rx, Cx)).Resize(tbl.Rows.Count, tbl.Columns.Count).Value = DataArray
        End If


        Return tbl
    End Function

    Function Excel_Cell(Row As Integer, Col As Integer, Optional ABSOLUTE As Integer = 0) As String
        Dim c As String
        c = Chr(((Col - 1) Mod 26) + 65)
        If Int((Col - 1) / 26) > 0 Then
            c = Chr(Int((Col - 1) / 26) + 64) & c
        End If

        Dim z1 As String = ""
        Dim z2 As String = ""

        If Row > 0 Then
            If ABSOLUTE = 1 Then
                z1 = "$"
            End If
            If ABSOLUTE = 2 Then
                z2 = "$"
            End If
            If ABSOLUTE = 3 Then
                z1 = "$"
                z2 = "$"
            End If
        End If

        Excel_Cell = z1 & c & z2 & IIf(Row > 0, CStr(Row), "")
    End Function
End Class