Imports Infragistics.Win.UltraWinGrid

Public Class SPRCDTL1
    Dim SPTCDTL1 As String
    Dim CMs As New Dictionary(Of String, CurrencyManager)
    Dim pymtColumnColors As New Dictionary(Of Integer, System.Drawing.Color) From {{1, System.Drawing.Color.MediumPurple}, {2, System.Drawing.Color.LightSeaGreen}, {3, System.Drawing.Color.PeachPuff}}
    Dim MAXCOLS As Integer = 20

    Dim wb As SpreadsheetGear.IWorkbook
    Dim ws As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing
    Dim rangeCopyFrom As SpreadsheetGear.IRange
    Dim rangePaste_To As SpreadsheetGear.IRange

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -120, 0, -18)
        Set_cmbYP("RYP1", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -120, 0, -12)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        MyBase.Get_SQL("*")

        Dim startPeriod As String = cmbRYP0.SelectedRow.Cells(0).Value
        Dim endPeriod As String = cmbRYP1.SelectedRow.Cells(0).Value

        Dim startDate As String = startPeriod.Substring(4, 2) & "/01/" & startPeriod.Substring(0, 4)
        Dim endDate As String = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(endPeriod.Substring(4, 2) & "/01/" & endPeriod.Substring(0, 4)))).ToString("MM/dd/yyyy")

        Dim sqlC As String = "(NVL(SPTCOOP1.QTY,0) * NVL(SPTCOOP1.VEHICLE_CPM,0) / 1000 + NVL(SPTCOOP1.OTHER_COST,0))"
        sql = "Select SPTCOOP1.*, SPTCOOP3.AUTH_LNO, " & vbCr
        sql &= $" Case When {sqlC} = 0 Then 0 Else ROUND(SPTCOOP1.OPEN_AMT * (SPTCOOP3.DIST_AMT / {sqlC}),2) END OPEN_AMT_COLLECTION," & vbCr
        sql &= $" Case When {sqlC} = 0 Then 0 Else ROUND(SPTCOOP1.PAID_AMT * (SPTCOOP3.DIST_AMT / {sqlC}),2) END PAID_AMT_COLLECTION," & vbCr
        sql &= " SPTCOOP3.ITEM_CODE, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, ICTCOLL1.BRAND_CODE" & vbCr
        sql &= " from SPTCOOP3" & sql_TABLE_NAMEs & vbCr

        'sql = "Select SPTCOOP1.*, SPTCOOP3.AUTH_LNO, " & vbCr
        'sql &= " ROUND(SPTCOOP1.OPEN_AMT * (SPTCOOP3.DIST_AMT / (NVL(SPTCOOP1.QTY,0) * NVL(SPTCOOP1.VEHICLE_CPM,0) / 1000 + NVL(SPTCOOP1.OTHER_COST,0))),2) OPEN_AMT_COLLECTION," & vbCr
        'sql &= " ROUND(SPTCOOP1.PAID_AMT * (SPTCOOP3.DIST_AMT / (NVL(SPTCOOP1.QTY,0) * NVL(SPTCOOP1.VEHICLE_CPM,0) / 1000 + NVL(SPTCOOP1.OTHER_COST,0))),2) PAID_AMT_COLLECTION," & vbCr
        'sql &= " SPTCOOP3.ITEM_CODE, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, ICTCOLL1.BRAND_CODE" & vbCr
        'sql &= " from SPTCOOP3" & sql_TABLE_NAMEs & vbCr

        Dim sql_filter As String = " and SPTCOOP1.OPS_YYYYPP between '" & startPeriod & "' and '" & endPeriod & "'" & vbCr

        Select Case optInclude.Value
            Case "O"
                sql_filter &= " and SPTCOOP1.OPEN_AMT <> 0" & vbCr

            Case "P"
                sql_filter &= " and SPTCOOP1.PAID_AMT <> 0" & vbCr

            Case Else
                ' sql_filter &= " and (NVL(SPTCOOP1.OPEN_AMT,0) <> 0 or NVL(SPTCOOP1.PAID_AMT,0) <> 0)" & vbCr
                ' the line above keeps promos from appearing on the report where the SPTCOOP3 records are zeroed out yet there is accrual or payment history

        End Select

        sql &= ASCMAIN1.SQL_Add_WHERE(sql_JOIN & sql_WHERE & sql_filter)

        SPTCDTL1 = ASCMAIN1.Temp_Table(sql)

        Dim sqlSPTCOOPX = "select * from " & SPTCDTL1
        If Not dst.Tables.Contains("SPTCOOPX") Then
            Create_TDA(dst.Tables.Add, "SPTCOOPX", sqlSPTCOOPX, 0, False, "", 0)
        End If
        Fill_Records("SPTCOOPX", String.Empty, True, sqlSPTCOOPX)

        Dim sqlDetailKey As String = "SPTCOOP3.AUTH_LNO, SPTCOOP3.AUTH_NO || '_' || SPTCOOP3.COLLECTION_CODE DETAIL_KEY, "
        Dim sqlDetails As String = Replace(sql, "SPTCOOP3.AUTH_LNO,", sqlDetailKey)
        If Not dst.Tables.Contains("SPTCOOPY") Then
            Create_TDA(dst.Tables.Add, "SPTCOOPY", sqlDetails, 0, False, "", 0)
            dst.Tables("SPTCOOPY").Columns.Add("PYMT_COUNT", GetType(System.String))
            For i As Integer = 1 To MAXCOLS
                dst.Tables("SPTCOOPY").Columns.Add("VEND_CODE_PYMT_" & i.ToString, GetType(System.String))
                dst.Tables("SPTCOOPY").Columns.Add("PYMT_REF_NO_PYMT_" & i.ToString, GetType(System.String))
                dst.Tables("SPTCOOPY").Columns.Add("DIST_AMT_PYMT_" & i.ToString, GetType(System.Decimal))
            Next
        End If

        Dim sqlExtendedDetails As String = "select SPTPYMT2.*, SPTPYMT1.PYMT_REF_NO, SPTPYMT1.PYMT_REF_AMT PYMT_REF_AMT2, SPTPYMT1.PYMT_TYPE, SPTPYMT1.CUST_CODE, SPTPYMT1.VEND_CODE" & vbCrLf _
            & "  , SPTPYMT1.NOTES_PYMT, SPTPYMT1.PYMT_REF_DATE, SPTPYMT1.PYMT_CTL_NO" & vbCrLf _
            & "  , SPTPYMT3.AUTH_LNO, SPTPYMT3.COLLECTION_CODE, SPTPYMT3.AUTH_NO || '_' || SPTPYMT3.COLLECTION_CODE DETAIL_KEY, SPTPYMT3.DIST_AMT_PYMT" & vbCrLf _
            & "          from SPTPYMT2, SPTPYMT1, SPTPYMT3, " & SPTCDTL1 & " SPTCDTL1 " & vbCrLf _
            & "  WHERE SPTPYMT1.PYMT_NO = SPTPYMT2.PYMT_NO" & vbCrLf _
            & "  AND SPTPYMT2.PYMT_NO = SPTPYMT3.PYMT_NO" & vbCrLf _
            & "  AND SPTPYMT2.PYMT_LNO = SPTPYMT3.PYMT_LNO" & vbCrLf _
            & "  AND SPTPYMT2.AUTH_NO = SPTPYMT3.AUTH_NO" & vbCrLf _
            & "  AND SPTPYMT2.AUTH_PNO = SPTPYMT3.AUTH_PNO" & vbCrLf _
            & " and SPTPYMT3.AUTH_NO = SPTCDTL1.AUTH_NO  anD SPTPYMT3.COLLECTION_CODE = SPTCDTL1.COLLECTION_CODE"

        If Not dst.Tables.Contains("SPTCOOPZ") Then
            Create_TDA(dst.Tables.Add, "SPTCOOPZ", sqlExtendedDetails, 0, False, "", 0)
        End If

        'Create_Relation("SPTCOOPY", "SPTCOOPZ", "DETAIL_KEY")

        If chkExtended.Checked Then
            Fill_Records("SPTCOOPY", String.Empty, True, sqlDetails)
            Fill_Records("SPTCOOPZ", String.Empty, True, sqlExtendedDetails)
            For Each rowSPTCOOPY As DataRow In dst.Tables("SPTCOOPY").Select("", "AUTH_NO")
                Populate_Extended_Detaild(rowSPTCOOPY)
            Next
        End If

        sql = "Select " & sql_SELECT_cols.Replace("SPTCOOP1.", "").Replace("SPTCOOP3.", "").Replace("ICTCOLL1.", "") & vbCr
        sql &= ", AUTH_NO, AUTH_LNO, 0 TOTAL "
        sql &= " FROM " & SPTCDTL1

        sql = "Insert Into " & ASTSRPT1 & " " & sql
        ASCDATA1.ExecuteSQL(sql)


        Create_XLS()

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Sub Create_XLS()

        Dim sqlAUTH_NO As String = $"Select Distinct AUTH_NO from {SPTCDTL1} SPTCDTL1"
        Dim sqlA As String = ""
        Dim sqlP As String = ""
        Dim sqlPxx As String = ""
        Dim RYPX As String = RYP0
        Dim PX As Integer = 0
        Dim RYP_OPENING As String = ASCMAIN1.Period_Calc(RYP0, -1)


        Dim YPsList As New List(Of String)
        YPsList.Add("")
        Dim YPsLegends As New Dictionary(Of String, String)
        YPsLegends.Add("", "")
        Dim sqlSet As String = ""
        Do
            YPsList.Add(RYPX)
            YPsLegends.Add(RYPX, Mid(ASCMAIN1.Get_Legend(RYPX), 10, 6))
            PX += 1
            sqlA &= $", SUM (Case When GLTTSPCA.OPS_YYYYPP = '{RYPX}' and GLTTSPCA.OPS_YYYYPP <> '{ASCMAIN1.CYP}' Then GLTTSPCA.DIST_AMT Else 0 End) A{Format(PX, "00")}"
            sqlP &= $", SUM (Case When SPTPYMT1.OPS_YYYYPP = '{RYPX}' Then SPTPYMT2.PYMT_REF_AMT Else 0 End) P{Format(PX, "00")}" & vbCrLf
            sqlPxx &= $", TABLE.COLUMN{Format(PX, "00")}"

            If RYPX <> ASCMAIN1.CYP Then
                sqlSet &= $", A{Format(PX, "00")} = NVL(A{Format(PX, "00")},0) + NVL(P{Format(PX, "00")},0)"
            End If

            RYPX = ASCMAIN1.Period_Calc(RYPX, 1)
        Loop Until RYPX > RYP1

        Dim YPs() As String = YPsList.ToArray

        sqlP = "Select SPTPYMT2.AUTH_NO" & vbCrLf & sqlP & $" from SPTPYMT1,SPTPYMT2
where SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO
  and SPTPYMT1.OPS_YYYYPP BETWEEN '{RYP0}' AND '{RYP1}'
  and SPTPYMT2.AUTH_NO in ({sqlAUTH_NO})
group by SPTPYMT2.AUTH_NO"

        sqlA = "Select GLTTSPCA.DIST_CTL_NO AUTH_NO" & vbCrLf & sqlA & $" from GLTTSPCA
where GLTTSPCA.OPS_YYYYPP BETWEEN '{RYP0}' AND '{RYP1}'
  and GLTTSPCA.DIST_CTL_NO in ({sqlAUTH_NO})
group by GLTTSPCA.DIST_CTL_NO"

        ASCMAIN1.sql = $"
Select CUST_CODE, EXPENSE_TYPE_CODE, SPTCOOP1.AUTH_NO, AUTH_DATE, SREP_CODE, CUST_REF_NUM 
, BOOKING_NAME, SEASON_CODE, EVENT_TYPE_CODE, OPS_YYYYPP, VEHICLE_CODE, DATE_START, DATE_END
, STATUS_CODE, VERIFIED_AS_OPEN_BY, VERIFIED_AS_OPEN_DATE
, OPEN_AMT DIST_AMT_ORIG
, OPEN_AMT DIST_AMT_ADJ
, OPEN_AMT, PAID_AMT, PYMTS
, OPENING.BEGBAL
{Replace(sqlPxx, "TABLE.COLUMN", "A.A")}
{Replace(sqlPxx, "TABLE.COLUMN", "P.P")}
from SPTCOOP1
, ({sqlA}) A
, ({sqlP}) P
, (Select DETL_CTL_NO AUTH_NO, Sum (CREC_AMT) BEGBAL from GLTCREC3 
where CREC_TYPE_CODE = 'PX' and OPS_YYYYPP = '{RYP_OPENING}'
  and CREC_AMT <> 0 and DETL_CTL_NO in ({sqlAUTH_NO}) group by DETL_CTL_NO) OPENING
where SPTCOOP1.AUTH_NO IN ({sqlAUTH_NO})
  and A.AUTH_NO (+) = SPTCOOP1.AUTH_NO
  and P.AUTH_NO (+) = SPTCOOP1.AUTH_NO
  and OPENING.AUTH_NO (+) = SPTCOOP1.AUTH_NO"

        Dim SPTCDTLX As String = ASCMAIN1.Temp_Table()

        ASCMAIN1.sql = $"Update {SPTCDTLX} Set " & Mid(sqlSet, 3)
        ASCDATA1.ExecuteSQL()
        
        ASCMAIN1.sql = $"Update {SPTCDTLX} Set DIST_AMT_ORIG = 0, DIST_AMT_ADJ = 0"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = $"Select AUTH_NO, OPS_YYYYPP, SUM (DIST_AMT) DIST_AMT from (
Select GLTTSPCA.DIST_CTL_NO AUTH_NO, 'A' PROMO_TYPE, GLTTSPCA.OPS_YYYYPP, SUM (GLTTSPCA.DIST_AMT) DIST_AMT
from GLTTSPCA
where GLTTSPCA.OPS_YYYYPP BETWEEN '{RYP0}' AND '{RYP1}'
  and GLTTSPCA.DIST_CTL_NO in ({sqlAUTH_NO})
and GLTTSPCA.DIST_AMT <> 0
group by GLTTSPCA.DIST_CTL_NO, GLTTSPCA.OPS_YYYYPP
union
Select SPTPYMT2.AUTH_NO, 'P' PROMO_TYPE, SPTPYMT1.OPS_YYYYPP, SUM (SPTPYMT2.PYMT_REF_AMT) DIST_AMT
 from SPTPYMT1,SPTPYMT2
where SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO
  and SPTPYMT1.OPS_YYYYPP BETWEEN '{RYP0}' AND '{RYP1}'
  and SPTPYMT2.AUTH_NO in ({sqlAUTH_NO})
and SPTPYMT2.PYMT_REF_AMT <> 0
group by SPTPYMT2.AUTH_NO, SPTPYMT1.OPS_YYYYPP
  ) group by AUTH_NO, OPS_YYYYPP
"
        Dim SPTCDTLY As String = ASCMAIN1.Temp_Table()



        ASCMAIN1.sql = $"Select AUTH_NO, MIN (OPS_YYYYPP) MINYP, SUM (DIST_AMT) DIST_AMT from {SPTCDTLY} SPTCDTLY group by AUTH_NO"
        Dim SPTCDTLZ As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL($"Alter Table {SPTCDTLZ} Add Primary Key (AUTH_NO)")
        ASCDATA1.ExecuteSQL($"Alter Table {SPTCDTLZ} Add DIST_AMT_ORIG NUMBER (13,2)")
        ASCDATA1.ExecuteSQL($"Update {SPTCDTLZ} SPTCDTLZ Set DIST_AMT_ORIG = (Select DIST_AMT from {SPTCDTLY} SPTCDTLY where AUTH_NO = SPTCDTLZ.AUTH_NO and OPS_YYYYPP = SPTCDTLZ.MINYP) where MINYP is Not Null")
        ASCDATA1.ExecuteSQL($"Alter Table {SPTCDTLZ} Add DIST_AMT_ADJ NUMBER (13,2)")
        ASCDATA1.ExecuteSQL($"Update {SPTCDTLZ} SPTCDTLZ Set DIST_AMT_ADJ = (Select SUM (DIST_AMT) from {SPTCDTLY} SPTCDTLY where AUTH_NO = SPTCDTLZ.AUTH_NO and OPS_YYYYPP <> SPTCDTLZ.MINYP) where MINYP is Not Null")


        '        ASCMAIN1.sql = $"Select SPTCDTLX.*, SPTCDTLZ.MINYP, SPTCDTLZ.DIST_AMT_ORIG, SPTCDTLZ.DIST_AMT_ADJ
        'From {SPTCDTLX} SPTCDTLX, {SPTCDTLZ} SPTCDTLZ where SPTCDTLZ.AUTH_NO (+) = SPTCDTLX.AUTH_NO"

        ASCDATA1.ExecuteSQL($"Update {SPTCDTLX} SPTCDTLX Set DIST_AMT_ORIG = (Select DIST_AMT_ORIG from {SPTCDTLZ} SPTCDTLZ where AUTH_NO = SPTCDTLX.AUTH_NO)")
        ASCDATA1.ExecuteSQL($"Update {SPTCDTLX} SPTCDTLX Set DIST_AMT_ADJ  = (Select DIST_AMT_ADJ  from {SPTCDTLZ} SPTCDTLZ where AUTH_NO = SPTCDTLX.AUTH_NO)")


        ASCMAIN1.sql = $"Select SPTCDTLX.* from {SPTCDTLX} SPTCDTLX"
        Dim DataTable As DataTable = ASCDATA1.GetDataTable

        Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & $"Promo_Summary_{XNO}.xlsx"

        wb = SpreadsheetGear.Factory.GetWorkbook()
        ws = wb.Worksheets(0)
        ws.Name = "Promo Summary"

        Dim r0 As Int32 = 1 ' Row containing headings
        Dim rTotal As Int32 = DataTable.Select("").Length
        Dim cTotal As Int32 = DataTable.Columns.Count
        Dim r As Int32 = 0

        Dim c_BEG As Integer = -1
        Dim c_ACC As Integer = -1
        Dim c_PMT As Integer = -1
        Dim c_END As Integer = -1

        Dim c As Integer = -1

        c += 1 : Format_Column(r0, c, "Customer", 12, "@", "CUST_CODE")
        c += 1 : Format_Column(r0, c, "Expense Type", 10, "@", "EXPENSE_TYPE_CODE")
        c += 1 : Format_Column(r0, c, "Auth No", 10, "@", "AUTH_NO")
        c += 1 : Format_Column(r0, c, "Auth Date", 12, "MM/DD/YY", "AUTH_DATE")
        c += 1 : Format_Column(r0, c, "SRep", 7, "@", "SREP_CODE")
        c += 1 : Format_Column(r0, c, "Cust Ref", 10, "@", "CUST_REF_NUM")
        c += 1 : Format_Column(r0, c, "Booking", 20, "@", "BOOKING_NAME")
        c += 1 : Format_Column(r0, c, "Season", 8, "@", "SEASON_CODE")
        c += 1 : Format_Column(r0, c, "Event Type", 12, "@", "EVENT_TYPE_CODE")
        c += 1 : Format_Column(r0, c, "YYYYPP", 10, "@","OPS_YYYYPP")
        c += 1 : Format_Column(r0, c, "Vehicle", 12, "@","VEHICLE_CODE")
        c += 1 : Format_Column(r0, c, "Date Start", 12, "MM/DD/YY","DATE_START")
        c += 1 : Format_Column(r0, c, "Date End", 12, "MM/DD/YY","DATE_END")
        c += 1 : Format_Column(r0, c, "Status", 10, "@","STATUS_CODE")
        c += 1 : Format_Column(r0, c, "Ver by", 12, "@", "VERIFIED_AS_OPEN_BY")
        c += 1 : Format_Column(r0, c, "Ver Date", 12, "MM/DD/YY", "VERIFIED_AS_OPEN_DATE")

        c += 1 : Format_Column(r0, c, "Orig Amt", 12, "#,###", "DIST_AMT_ORIG", rTotal)
        c += 1 : Format_Column(r0, c, "Adjustments", 12, "#,###", "DIST_AMT_ADJ", rTotal)

        c += 1 : Format_Column(r0, c, "Open", 12, "#,###", "OPEN_AMT", rTotal)
        c += 1 : Format_Column(r0, c, "Paid", 12, "#,###", "PAID_AMT", rTotal)
        c += 1 : Format_Column(r0, c, "Pymts", 6, "#,###", "PYMTS")

        ws.Cells(r0, 0, r0, c).Interior.Color = SpreadsheetGear.Colors.LightPink

        c += 1 : Format_Column(r0, c, "Beg Bal", 12, "#,###", "BEGBAL", rTotal)
        'ws.Cells(r0 - 1, c).Formula = $"=subtotal(9,{Excel_Cell0(r0 + 1, c)}:{Excel_Cell0(r0 + rTotal, c)})"
        ws.Cells(r0, c, r0, c).Interior.Color = SpreadsheetGear.Colors.LightGoldenrodYellow
        c_BEG = c

        c_ACC = c + 1
        For pxi As Integer = 1 To PX
            Dim HDG As String = $"Acc {YPsLegends(YPs(pxi))}"
            c += 1 : Format_Column(r0, c, HDG, 12, "#,###", $"A{Format(pxi, "00")}", rTotal)
            'ws.Cells(r0 - 1, c).Formula = $"=subtotal(9,{Excel_Cell0(r0 + 1, c)}:{Excel_Cell0(r0 + rTotal, c)})"
        Next
        ws.Cells(r0, c - PX + 1, r0, c).Interior.Color = SpreadsheetGear.Colors.PaleGreen

        c_PMT = c + 1
        For pxi As Integer = 1 To PX
            Dim HDG As String = $"Pmt { YPsLegends(YPs(pxi))}"
            c += 1 : Format_Column(r0, c, HDG, 12, "#,###", $"P{Format(pxi, "00")}", rTotal)
            'ws.Cells(r0 - 1, c).Formula = $"=subtotal(9,{Excel_Cell0(r0 + 1, c)}:{Excel_Cell0(r0 + rTotal, c)})"
        Next
        ws.Cells(r0, c - PX + 1, r0, c).Interior.Color = SpreadsheetGear.Colors.PaleTurquoise

        c += 1 : Format_Column(r0, c, "End Bal", 12, "#,###", "ENDBAL", rTotal)
        ws.Cells(r0, c, r0, c).Interior.Color = SpreadsheetGear.Colors.LightGoldenrodYellow
        c_END = c

        Dim FX As String = $"={Excel_Cell0(r0 + 1, c_BEG)} + SUM({Excel_Cell0(r0 + 1, c_ACC)}:{Excel_Cell0(r0 + 1, c_PMT - 1)}) - +SUM({Excel_Cell0(r0 + 1, c_PMT)}:{Excel_Cell0(r0 + 1, c_END - 1)})"
        ws.Cells(r0 + 1, c).Formula = FX

        rangeCopyFrom = ws.Cells(r0 + 1, c)
        rangePaste_To = ws.Range(r0 + 1, c, r0 + rTotal, c)
        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)


        ws.Cells(r0, 0, r0, c).AutoFilter()

        ws.Cells(r0 + 1, 2).Activate()
        ws.WindowInfo.FreezePanes = True

        ws.Range(r0 + 1, 0).CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)


        wb.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Show_Document(XLS_FILENAME)
    End Sub

    Sub Format_Column(r0 As Integer, c As Integer, Caption As String, Width As Integer, Format As String, COLUMN_NAME As String, Optional rTotal As Int32 = -1)

        With ws.Cells(r0, c)
            .EntireColumn.NumberFormat = Format
            .EntireColumn.ColumnWidth = Width
            .EntireColumn.Locked = True
            If Format = "MM/DD/YY" Then
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If
            .Value = Caption
            If rTotal <> -1 Then
                ws.Cells(r0 - 1, c).Formula = $"=subtotal(9,{Excel_Cell0(r0 + 1, c)}:{Excel_Cell0(r0 + rTotal, c)})"
            End If
        End With

    End Sub


    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)

        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
        grdASTEXPT1.DataSource = dst.Tables("SPTCOOPX")
        grdASTEXPT1.Text = "Promo Expense Detail - " & RYPLEGEND0 & " thru " & RYPLEGEND1
        grdASTEXPT1.Name = "grdASTEXPT1"
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text
        Format_DX_Grid(grdASTEXPT1)

        If chkExtended.Checked Then
            Dim tbl As DataTable = dst.Tables("SPTCOOPY")
            Dim grdASTEXPT2 As UltraGrid = Add_DX_Grid(tabDataExports, "Promo Expense Detail Extended - " & RYPLEGEND0 & " thru " & RYPLEGEND1, "grdASTEXPT2")
            'grdASTEXPT2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
            grdASTEXPT2.DataMember = tbl.TableName
            grdASTEXPT2.DataSource = dst
            'grdASTEXPT1.DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.Horizontal
            Format_DX_Grid(grdASTEXPT2)
        End If

        UltraTabControl1.Tabs("Data Exports").Visible = True

    End Sub

    Sub Format_DX_Grid(grd As UltraWinGrid.UltraGrid)

        Set_DX_Column(grd, "")

        Set_DX_Column(grd, "CUST_CODE", "Customer", 90, , , System.Drawing.Color.Gold)
        Set_DX_Column(grd, "EXPENSE_TYPE_CODE", "Expense Type", 50)
        Set_DX_Column(grd, "AUTH_NO", "Auth No", 70)
        Set_DX_Column(grd, "AUTH_DATE", "Auth Date", 100)
        Set_DX_Column(grd, "SREP_CODE", "SRep", 50)
        Set_DX_Column(grd, "CUST_REF_NUM", "Cust Ref", 50)
        Set_DX_Column(grd, "BOOKING_NAME", "Booking", 50)
        Set_DX_Column(grd, "SEASON_CODE", "Season", 50)
        Set_DX_Column(grd, "EVENT_TYPE_CODE", "Event Type", 50)
        Set_DX_Column(grd, "OPS_YYYYPP", "YYYYPP", 50)
        Set_DX_Column(grd, "VEHICLE_CODE", "Vehicle", 50)
        Set_DX_Column(grd, "DATE_START", "Date Start", 100)
        Set_DX_Column(grd, "DATE_END", "Date End", 100)
        Set_DX_Column(grd, "STATUS_CODE", "Status", 50)
        Set_DX_Column(grd, "COLLECTION_CODE", "Collec", 70)
        Set_DX_Column(grd, "BRAND_CODE", "Brand", 50)
        Set_DX_Column(grd, "DIST_AMT", "Dist Amt", 50)

        Set_DX_Column(grd, "VERIFIED_AS_OPEN_BY", "Ver by", 70, , , System.Drawing.Color.Orange)
        Set_DX_Column(grd, "VERIFIED_AS_OPEN_DATE", "Ver Date", 100, , , System.Drawing.Color.Orange)

        Set_DX_Column(grd, "OPEN_AMT_COLLECTION", "Open Amt", 90, "#,##0", , System.Drawing.Color.LightBlue)
        Set_DX_Column(grd, "PAID_AMT_COLLECTION", "Paid Amt", 90, "#,##0", , System.Drawing.Color.LightGreen)

        Create_Summary(grd, "OPEN_AMT_COLLECTION")
        Create_Summary(grd, "PAID_AMT_COLLECTION")

        grd.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True

        Sort_grdColumns(grd, "CUST_CODE")

        If grd.Name = "grdASTEXPT2" Then

            Set_DX_Column(grd, "PYMT_COUNT", "# Pymts", 40)

            For i As Integer = 1 To MAXCOLS
                Dim VEND_CODE_PYMT_COL As String = "VEND_CODE_PYMT_" & i
                Dim PYMT_REF_NO_PYMT_COL As String = "PYMT_REF_NO_PYMT_" & i
                Dim PYMT_REF_AMT_PYMT_COL As String = "DIST_AMT_PYMT_" & i

                Set_DX_Column(grd, VEND_CODE_PYMT_COL, "Vendor Name", 100, , , pymtColumnColors(1 + (i - 1) Mod 3))
                Set_DX_Column(grd, PYMT_REF_NO_PYMT_COL, "Invoice #", 100, , , pymtColumnColors(1 + (i - 1) Mod 3))
                Set_DX_Column(grd, PYMT_REF_AMT_PYMT_COL, "Amount", 90, "#,##0.00", , pymtColumnColors(1 + (i - 1) Mod 3))

                Create_Summary(grd, PYMT_REF_AMT_PYMT_COL)
            Next

            'Set_DX_Column(grd, "", , , , , , 1)
            'Set_DX_Column(grd, "PYMT_NO", "Pymt No", 90, , , System.Drawing.Color.Aquamarine, 1)
            'Set_DX_Column(grd, "VEND_CODE", "Vendor", 100, , , , 1)
            'Set_DX_Column(grd, "PYMT_NO", "Pymt Lno", 50, , , , 1)
            'Set_DX_Column(grd, "PYMT_REF_AMT", "Amount", 75, "#,##0", , System.Drawing.Color.LightSteelBlue, 1)
            'Set_DX_Column(grd, "CLOSED", "Closed", 50, , , , 1)
            'Set_DX_Column(grd, "NOTES_PYMT", "Notes", 100, , , , 1)
            'Set_DX_Column(grd, "PYMT_TYPE", "Type", 75, , , , 1)
            'Set_DX_Column(grd, "PYMT_REF_NO", "Reference", 100, , , , 1)
            'Set_DX_Column(grd, "PYMT_REF_DATE", "Date", 75, , , , 1)
            'Set_DX_Column(grd, "PYMT_REF_AMT2", "Total Pymt", 75, "#,##0", , System.Drawing.Color.LightSeaGreen, 1)
            'Set_DX_Column(grd, "PYMT_CTL_NO", "Pymt Ctl No", 100, , , , 1)
            'Set_DX_Column(grd, "COLLECTION_CODE", "Collection", 80, , , , 1)
            'Set_DX_Column(grd, "NOTES_PYMT_APPR", "Appr Notes", 100, , , , 1)
            'Set_DX_Column(grd, "CUST_CODE", "Customer", 80, , , , 1)

            'Create_Summary(grd, "PYMT_REF_AMT", , "SPTCOOPY_SPTCOOPZ")
            'Create_Summary(grd, "PYMT_REF_AMT2", , "SPTCOOPY_SPTCOOPZ")

            'grd.DisplayLayout.Bands(1).Columns("PYMT_NO").Header.Fixed = True

            'Sort_grdColumns(grd, "PYMT_NO", , 1)

        End If

    End Sub

    Sub Populate_Extended_Detaild(detailRow As DataRow)

        Dim DETAIL_KEY As String = detailRow.Item("DETAIL_KEY")
        Dim pymtCount As Integer = 0
        Dim MAXMAX As Integer = 0

        For Each rowSPTCOOPZ As DataRow In dst.Tables("SPTCOOPZ").Select("DETAIL_KEY = '" & DETAIL_KEY & "'", "PYMT_NO")
            pymtCount += 1
            Dim VEND_CODE As String = rowSPTCOOPZ.Item("VEND_CODE") & ""
            Dim PYMT_REF_NO As String = rowSPTCOOPZ.Item("PYMT_REF_NO") & ""
            Dim DIST_AMT_PYMT As Decimal = Val(rowSPTCOOPZ.Item("DIST_AMT_PYMT") & "")

            Dim detailColSuf As String = IIf(pymtCount >= MAXCOLS, $"_{MAXCOLS}", "_" & pymtCount)

            Dim VEND_CODE_PYMT_COL As String = "VEND_CODE_PYMT" & detailColSuf
            Dim PYMT_REF_NO_PYMT_COL As String = "PYMT_REF_NO_PYMT" & detailColSuf
            Dim DIST_AMT_PYMT_COL As String = "DIST_AMT_PYMT" & detailColSuf

            Dim VEND_CODE_PYMT As String = detailRow.Item(VEND_CODE_PYMT_COL) & ""
            Dim PYMT_REF_NO_PYMT As String = detailRow.Item(PYMT_REF_NO_PYMT_COL) & ""
            Dim DIST_AMT_PYMT_CURR As Decimal = Val(detailRow.Item(DIST_AMT_PYMT_COL) & "")

            detailRow.Item("PYMT_COUNT") = IIf(pymtCount > MAXCOLS, pymtCount, "")

            If pymtCount > MAXCOLS Then
                If pymtCount > MAXMAX Then MAXMAX = pymtCount

                Dim pymtVendors As String() = VEND_CODE_PYMT.Split(New Char() {","c})

                If Not pymtVendors.Contains(VEND_CODE) Then
                    detailRow.Item(VEND_CODE_PYMT_COL) = VEND_CODE_PYMT & "," & VEND_CODE
                End If

                detailRow.Item(PYMT_REF_NO_PYMT_COL) = PYMT_REF_NO_PYMT & "," & PYMT_REF_NO
                detailRow.Item(DIST_AMT_PYMT_COL) = DIST_AMT_PYMT + DIST_AMT_PYMT_CURR
            Else
                detailRow.Item(VEND_CODE_PYMT_COL) = VEND_CODE
                detailRow.Item(PYMT_REF_NO_PYMT_COL) = PYMT_REF_NO
                detailRow.Item(DIST_AMT_PYMT_COL) = DIST_AMT_PYMT
            End If

        Next

        If ASCMAIN1.Running_in_VS Then
            ' If MAXMAX > 9 Then MsgBox($"Max Pymts {MAXMAX}")
        End If
    End Sub
    Overrides Sub Verify_Special(ByVal eItemKey As String)

    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

End Class