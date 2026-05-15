Imports System.Math

Public Class ICRVDEF1

    Shadows SUBT As String = ""

    Dim ICTVDEF1 As String
    Dim ICTVDEF2 As String
    Dim ICTIVAR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("ICTPARM1")
        Call Get_PARM("GLTPARM1")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        SUBT = ""

        Prepare_dst(True)

        Check_if_Empty("ICTVDEF1")
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("SUMMARY", "1")
        SUBT = "Summary for Month Ending " & RYPLEGEND
        Generate_Report(RPT, , SUBT, "{ICTVDEF1.VARIANCE} <> 0.00")

        CR_params.Add("SUMMARY", "0")
        SUBT = "Cost Lot Detail for Month Ending " & RYPLEGEND
        Generate_Report(RPT, , SUBT, "{ICTVDEF1.VARIANCE} <> 0.00")

        tabData.Visible = True

        grdICTVDEF3.Visible = True
        grdICTVDEF3.Text = "Aged Valuation Summary by Item - " & RYPLEGEND
        grdICTVDEF3.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdICTVDEF3.DisplayLayout.Bands(0).SortedColumns.Add("COST_CATGY_CODE", False, True)
        grdICTVDEF3.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)

        Print_GL()

        Prepare_Data_Extracts()
        Prepare_Pivot_Tables()

    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
        '  Dim TBL As DataTable = dst.Tables("ICTVDEF1").Clone
        grdASTEXPT1.DataSource = dst.Tables("ICTVDEF1")
        grdASTEXPT1.Text = "Variances by Item - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 100)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70)
        ' Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY", "MB", 30)
        Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "Cost Catgy", 70)
        Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_VCOST", "Std VCOST", 90, "#.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MATLS", "Std MATLS", 90, "#.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOTAL", "Std Cost", 90, "#.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "QTY_USED", "FIFO Qty", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXT_STD_VCOST", "Ext VCOST", 90, "#,##0.00", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "EXT_STD_MATLS", "Ext MATLS", 90, "#,##0.00", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "EXT_STD_LANDG", "Ext LANDG", 90, "#,##0.00", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "EXT_STD_TOOLG", "Ext Tariff", 90, "#,##0.00", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "QTY_ONH", "OnH@EOM", 90, "#,##0", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "EXT_STD", "Ext Std", 90, "#,##0.00", "Sum", Color.Gold)
        Set_DX_Column(grdASTEXPT1, "VAR_VCOST", "Var VCOST", 90, "#,##0.00", "Sum", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VAR_MATLS", "Var MATLS", 90, "#,##0.00", "Sum", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VAR_LANDG", "Var LANDG", 90, "#,##0.00", "Sum", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VAR_TOOLG", "Var Tariff", 90, "#,##0.00", "Sum", Color.LightBlue)

        Set_DX_Column(grdASTEXPT1, "VARIANCE", "Total Var", 90, "#,##0.00", "Sum", Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "EXT_USED_VCOST", "Ext VCOST FIFO", 90, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXT_USED_MATLS", "Ext MATLS FIFO", 90, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXT_USED_LANDG", "Ext LANDG FIFO", 90, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXT_USED_TOOLG", "Ext Tariff FIFO", 90, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXT_USED", "Ext FIFO", 90, "#,##0.00", "Sum", Color.LightGreen)

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub

    Sub Prepare_Pivot_Tables()

        Dim XCT As Integer = 58 ' the 1-based Column where the Formulae cells begin
        ' previously I documented this (incorectly) as Number of Data Columns before the Formulas Columns
        Dim XCT_col As String = Excel_Cell(-1, XCT) ' "BF" ' Column where Formulae cells begin

        Dim FILENAME_SSG As String = CopyDataToSSG("ICTVDEF2", XCT - 1, "ITEM_CODE, RECEIPT_DATE DESC, RECEIPT_NO DESC")

        ASCMAIN1.Progress("Now Preparing Pivot Table")

        Dim xlsFilename As String = "ICRVDEF1.xlsx" ' this is the Template with the Pivot Table(s) defined
        Dim excel As New Microsoft.Office.Interop.Excel.Application
        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.Folders("SharedRoot") = "C:\SHARE\INT\"
        End If
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(ASCMAIN1.Folders("SharedRoot") & "Templates\" & xlsFilename)

        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets("Data")
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        Dim r As Integer = 0 ' row count for Data
        Dim iRow = 3 ' 1-based row where Data starts

        ' this is the old fashioned way to copy data from ado.net into interop XLS
        'If dst.Tables.Contains("ICTVDEF2") Then
        '    For Each rowICTVDEF2 As DataRow In dst.Tables("ICTVDEF2").Select("")
        '        For C As Integer = 1 To XCT - 1
        '            ws.Cells(iRow + r, C).value = rowICTVDEF2.Item(C - 1)
        '        Next
        '        r += 1
        '        If r Mod 100 = 0 Then
        '            ASCMAIN1.Progress("-", r)
        '        End If
        '        If r > 500 Then Exit For
        '    Next
        'End If

        r = dst.Tables("ICTVDEF2").Rows.Count

        ' Copy data from SSG
        Dim wbSSG As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME_SSG)
        Dim wsSSG As Microsoft.Office.Interop.Excel.Worksheet = wbSSG.Worksheets("Data")
        xlSourceRange = wsSSG.Range(Excel_Cell(1 + 1, 0 + 1) & ":" & Excel_Cell(1 + r, 0 + XCT - 1))
        xlDestRange = ws.Range(Excel_Cell(iRow - 1 + 1, 0 + 1) & ":" & Excel_Cell(iRow + r, 0 + XCT - 1))
        'xlSourceRange.Copy(xlDestRange)
        xlSourceRange.Copy() ' copies to the clipboard so that we do not paste formatting
        xlDestRange.PasteSpecial(Paste:=Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues)


        ' Copy SubTotal formula across top row
        Dim ST As Integer = 36 ' the number of formula fields appended to the data set
        'ws.Range(Excel_Cell(iRow - 1 - 1, XCT)).Formula = $"=SUBTOTAL(9,{XCT_col}{CStr(iRow)}:{XCT_col}{CStr(iRow - 1 + r)})"
        'xlSourceRange = ws.Range(Excel_Cell(iRow - 2, XCT) & ":" & Excel_Cell(iRow - 2, XCT))
        'xlDestRange = ws.Range(Excel_Cell(iRow - 2, XCT) & ":" & Excel_Cell(iRow - 2, XCT + ST - 1))
        'xlSourceRange.Copy(xlDestRange)
        For c As Integer = 1 To XCT + ST - 1
            Dim ST_type As Integer = 9
            If c = 1 Then ST_type = 3
            Dim F As String = $"=SUBTOTAL({ST_type},{Excel_Cell(iRow, c)}:{Excel_Cell(iRow - 1 + r, c)})"
            If ws.Range(Excel_Cell(iRow - 1 - 1, c)).Value & "" <> "" Then
                ws.Range(Excel_Cell(iRow - 1 - 1, c)).Formula = F
            End If
        Next
        'ws.Range(Excel_Cell(iRow - 1 - 1, 1)).Formula = $"=SUBTOTAL(3,{Excel_Cell(iRow, 1)}:{Excel_Cell(iRow - 1 + r, 1)})"
        'xlSourceRange = ws.Range(Excel_Cell(iRow - 2, XCT) & ":" & Excel_Cell(iRow - 2, XCT))
        'xlDestRange = ws.Range(Excel_Cell(iRow - 2, XCT) & ":" & Excel_Cell(iRow - 2, XCT + ST - 1))
        'xlSourceRange.Copy(xlDestRange)


        'ws.Range(Excel_Cell(iRow - 1 - 1, 20)).Formula = $"=SUBTOTAL(9,T3:T{CStr(iRow - 1 + r)})"
        'xlSourceRange = ws.Range(Excel_Cell(iRow - 2, 20) & ":" & Excel_Cell(iRow - 2, 20))
        'xlDestRange = ws.Range(Excel_Cell(iRow - 2, 20) & ":" & Excel_Cell(iRow - 2, XCT - 1))
        'xlSourceRange.Copy(xlDestRange)



        ' Copy down formulae columns
        xlSourceRange = ws.Range(Excel_Cell(iRow - 1 + 1, XCT) & ":" & Excel_Cell(iRow - 1 + 1, XCT + ST - 1))
        xlDestRange = ws.Range(Excel_Cell(iRow - 1 + 1, XCT) & ":" & Excel_Cell(iRow - 1 + r, XCT + ST - 1))
        xlSourceRange.Copy(xlDestRange)

        'ws.Cells(1, 1).Value = Now
        'ws.Cells(1, 2).Value = "'" & YYYY

        wb.Names.Add("Receipts", "=Data!$A$" & CStr(iRow - 1 + 0) & ":$CO$" & CStr(iRow - 1 + r)) ' from heading row to last row of data

        ASCMAIN1.Progress("-", "Pivots")

        Try
            wb.Worksheets("PIVOT").Activate()
            'wb.ActiveSheet.PivotTables("Receipts").PivotCache.Refresh()
            ' WHY DOESN'T THIS WORK?
            ' ?wb.ActiveSheet.PivotTables(1).NAME
            ' "Receipts"
        Catch ex As Exception

        End Try

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "Deferred Variance Analysis"
                'XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                XLS_FILENAME &= "-" & XNO & "-" & Format(XLS_NO, "000") & ".xlsx"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        Show_Document(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

    End Sub

    Overrides Sub Update_Record()

        Update_Record_TDA("ICTIVAR1")

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from " & ICTIVAR1 & ";" _
            & " Begin " _
            & "  Update ICTIVAR1 Set PV_DEF = 0, MV_DEF = 0, FV_DEF = 0, TV_DEF = 0 where OPS_YYYYPP = '" & RYP & "';" _
            & "  For R1 in C1 Loop" _
            & "   Update ICTIVAR1 Set PV_DEF = -1 * NVL(R1.PV_DEF,0), MV_DEF = -1 * NVL(R1.MV_DEF,0)" _
            & ", FV_DEF = -1 * NVL(R1.FV_DEF,0), TV_DEF = -1 * NVL(R1.TV_DEF,0)" _
            & "    where ITEM_CODE = R1.ITEM_CODE And OPS_YYYYPP = R1.OPS_YYYYPP;" _
            & "   If SQL%NOTFOUND Then " _
            & "    Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, PV_DEF, MV_DEF, FV_DEF, TV_DEF) " _
            & "     Values (R1.ITEM_CODE, R1.OPS_YYYYPP,  -1 * R1.PV_DEF,  -1 * R1.MV_DEF,  -1 * R1.FV_DEF,  -1 * R1.TV_DEF);" _
            & "   End If;" _
            & "  End Loop; " _
            & " End; " _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Call GL_Update()
    End Sub


    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        ASCMAIN1.sql = "Select * from ICTIVAR1 where ROWNUM < 1"
        ICTIVAR1 = ASCMAIN1.Temp_Table
        Create_TDA(dst.Tables.Add("ICTIVAR1"), ICTIVAR1, "*")

        ASCMAIN1.sql = "Select * from ICTCOST1"
        Create_TDA(dst.Tables.Add, "ICTCOST1", "**", 0, False, "", 1)
        Fill_Records("ICTCOST1")

        Dim ICTCOSTA As String
        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql = "Select '" & RYP & "' OPS_YYYYPP, ICTCOSTC.* from ICTCOSTC"
        Else
            ASCMAIN1.sql = "Select * from ICTCOSTA where OPS_YYYYPP = '" & RYP & "'"
        End If
        ICTCOSTA = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTA & " Add Primary Key (OPS_YYYYPP, ITEM_CODE)")

        Dim FYP As String = "000000" ' 1st period that we should be taking APTINVH5 seriously
        If ASCMAIN1.CLIENT = "INT" Then FYP = "201601"

        ASCMAIN1.sql = "Select X.ITEM_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_VCOST" & vbCrLf _
            & ", (NVL(ICTCOSTA.ITEM_COST_MATLS,0) + NVL(ICTCOSTA.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)) ITEM_COST_MATLS" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_LANDG" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_TOOLG" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_TOTAL" & vbCrLf _
            & ", NVL(X.QTY_ONH,0) * NVL(ICTCOSTA.ITEM_COST_VCOST,0) EXT_STD_VCOST" & vbCrLf _
            & ", NVL(X.QTY_ONH,0) * (NVL(ICTCOSTA.ITEM_COST_MATLS,0) + NVL(ICTCOSTA.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)) EXT_STD_MATLS" & vbCrLf _
            & ", NVL(X.QTY_ONH,0) * NVL(ICTCOSTA.ITEM_COST_LANDG,0) EXT_STD_LANDG" & vbCrLf _
            & ", NVL(X.QTY_ONH,0) * NVL(ICTCOSTA.ITEM_COST_TOOLG,0) EXT_STD_TOOLG" & vbCrLf _
            & ", X.QTY_ONH, NVL(X.QTY_ONH,0) * (NVL(ICTCOSTA.ITEM_COST_VCOST,0) + NVL(ICTCOSTA.ITEM_COST_MATLS,0) + NVL(ICTCOSTA.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA.ITEM_COST_OVRHDI,0)) EXT_STD" & vbCrLf _
            & " from ICTITEM1," & ICTCOSTA & " ICTCOSTA, " & vbCrLf _
            & " (Select ITEM_CODE, SUM (WHSE_QTY_ON_HAND) QTY_ONH" & vbCrLf _
            & " from " & IIf(RYP = ASCMAIN1.CYP, "ICTSTAT2", "ICTSTAT5") & " ICTSTAT2 " & vbCrLf _
            & IIf(RYP = ASCMAIN1.CYP, "", " where ICTSTAT2.OPS_YYYYPP = '" & RYP & "'") & vbCrLf _
            & " group by ITEM_CODE) X" & vbCrLf _
            & " where ICTCOSTA.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and NVL(X.QTY_ONH,0) <> 0"
        ICTVDEF1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTVDEF1 & " Add Primary Key (ITEM_CODE)")

        Dim sqlCols As String = "RECEIPT_NO,RECEIPT_LNO,ITEM_CODE,QTY_REC,QTY_INV,ITEM_COST_STD,PO_COST,
PO_ORDER_NO,PO_ORDER_LNO,ITEM_UOM,OPS_YYYYPP,COST_CATGY_CODE,PROD_CODE,LOCATION_CODE,
BAR_CODE,VEND_WHSE_CODE,BM_ISSUE_SEL,BM_ISSUE_NO,REC_REF,EXT_COST_MATLS,TRAN_PV,
TRAN_MV,TRAN_CV,ACCRUAL_STATUS,AMT_INV,TRAN_FV,TRAN_TV,PO_COST_FRT,PO_COST_TRF"

        ASCMAIN1.sql = "Select ICTIREC2.*,ICTIREC1.SOURCE_DOC_NO,ICTIREC1.RECEIPT_DATE" & vbCrLf _
            & ", ICTIREC1.VEND_CODE" & vbCrLf _
            & ", APTACRCX.CTL_STATUS_LANDG, APTACRCX.VOUCHER_NO_LANDG" & vbCrLf _
            & ", APTACRCX.AMT_ACC_LANDG, 0 AMT_INV_LANDG, APTACRCX.AMT_USED_LANDG" & vbCrLf _
            & ", APTACRCX.CTL_STATUS_TOOLG, APTACRCX.VOUCHER_NO_TOOLG" & vbCrLf _
            & ", APTACRCX.AMT_ACC_TOOLG, 0 AMT_INV_TOOLG, APTACRCX.AMT_USED_TOOLG" & vbCrLf _
            & $" from (Select {sqlCols} from ICTIREC2) ICTIREC2,ICTIREC1," & ICTVDEF1 & " ICTVDEF1" & vbCrLf _
            & ", (Select APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO" & vbCrLf _
            & ", MAX (CASE WHEN APTACRC1.ACCRUAL_CODE = 'FRT' THEN CASE WHEN APTINVH1.VOUCHER_NO IS NULL THEN NULL ELSE APTACRC1.CTL_STATUS END ELSE NULL END) CTL_STATUS_LANDG" & vbCrLf _
            & ", MAX (CASE WHEN APTACRC1.ACCRUAL_CODE = 'FRT' THEN APTINVH1.VOUCHER_NO ELSE NULL END) VOUCHER_NO_LANDG" & vbCrLf _
            & ", SUM (CASE WHEN APTACRC1.ACCRUAL_CODE = 'FRT' THEN APTACRC1.COST_ACC ELSE NULL END) AMT_ACC_LANDG" & vbCrLf _
            & ", SUM (CASE WHEN APTACRC1.ACCRUAL_CODE = 'FRT' THEN CASE WHEN NVL(APTACRC1.CTL_STATUS,'0') = '1' AND APTINVH1.VOUCHER_NO IS NOT NULL THEN APTACRC1.COST_ACT ELSE APTACRC1.COST_ACC END ELSE NULL END) AMT_USED_LANDG" & vbCrLf _
            & ", MAX (CASE WHEN APTACRC1.ACCRUAL_CODE = 'TRF' THEN CASE WHEN APTINVH1.VOUCHER_NO IS NULL THEN NULL ELSE APTACRC1.CTL_STATUS END ELSE NULL END) CTL_STATUS_TOOLG" & vbCrLf _
            & ", MAX (CASE WHEN APTACRC1.ACCRUAL_CODE = 'TRF' THEN APTINVH1.VOUCHER_NO ELSE NULL END) VOUCHER_NO_TOOLG" & vbCrLf _
            & ", SUM (CASE WHEN APTACRC1.ACCRUAL_CODE = 'TRF' THEN APTACRC1.COST_ACC ELSE NULL END) AMT_ACC_TOOLG" & vbCrLf _
            & ", SUM (CASE WHEN APTACRC1.ACCRUAL_CODE = 'TRF' THEN CASE WHEN NVL(APTACRC1.CTL_STATUS,'0') = '1' AND APTINVH1.VOUCHER_NO IS NOT NULL THEN APTACRC1.COST_ACT ELSE APTACRC1.COST_ACC END ELSE NULL END) AMT_USED_TOOLG" & vbCrLf _
            & " from APTACRC1,APTINVH1" & vbCrLf _
            & " where CTL_TYPE in ('F','T')" & vbCrLf _
            & "   and CTL_DATE <= (SELECT PRD_END_DATE FROM GLTPARM2 WHERE OPS_YYYYPP = '" & RYP & "')" & vbCrLf _
            & "   and APTINVH1.VOUCHER_NO (+) = APTACRC1.VOUCHER_NO" & vbCrLf _
            & $"   and APTINVH1.OPS_YYYYPP (+) <= '{RYP}'" & vbCrLf _
            & " group by APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO) APTACRCX" & vbCrLf _
            & " where ICTIREC2.ITEM_CODE = ICTVDEF1.ITEM_CODE" & vbCrLf _
            & "   And ICTIREC2.QTY_REC <> 0" & vbCrLf _
            & "   And APTACRCX.RECEIPT_NO (+) = ICTIREC2.RECEIPT_NO" & vbCrLf _
            & "   And APTACRCX.RECEIPT_LNO (+) = ICTIREC2.RECEIPT_LNO" & vbCrLf _
            & "   And ICTIREC1.RECEIPT_DATE <= (SELECT PRD_END_DATE FROM GLTPARM2 WHERE OPS_YYYYPP = '" & RYP & "')" & vbCrLf _
            & "   and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO"
        ICTVDEF2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTVDEF2 & " Add Primary Key (RECEIPT_NO,RECEIPT_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTVDEF2 & " Add QTY_USED NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTVDEF2 & " Add AGE_DAYS NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTVDEF2 & "_1 on " & ICTVDEF2 & " (ITEM_CODE,RECEIPT_DATE)")


        ASCDATA1.ExecuteSQL("Update " & ICTVDEF2 & " Set QTY_INV =0, AMT_INV = 0")


        If ASCMAIN1.Running_in_VS Then
            Stop
            'ASCDATA1.ExecuteSQL("Delete from " & ICTVDEF1 & " where ITEM_CODE <> 'VAZV024'")
            'ASCDATA1.ExecuteSQL("Delete from " & ICTVDEF2 & " where ITEM_CODE <> 'VAZV024'")
            'ASCDATA1.ExecuteSQL("Delete from " & ICTVDEF1 & " where ITEM_CODE <> 'CC019A04'")
            'ASCDATA1.ExecuteSQL("Delete from " & ICTVDEF2 & " where ITEM_CODE <> 'CC019A04'")
        End If

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select APTINVH5.RECEIPT_NO, APTINVH5.RECEIPT_LNO" & vbCrLf _
            & " , Sum (NVL(APTINVH5.INV_QTY,0)) QTY_INV" & vbCrLf _
            & " , Sum (NVL(APTINVH5.INV_QTY,0) * NVL(APTINVH5.INV_COST,0)) AMT_INV" & vbCrLf _
            & "   from APTINVH5,APTINVH1," & ICTVDEF2 & " ICTVDEF2" & vbCrLf _
            & "   where APTINVH5.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
            & "     and APTINVH5.RECEIPT_NO = ICTVDEF2.RECEIPT_NO" & vbCrLf _
            & "     and APTINVH5.RECEIPT_LNO = ICTVDEF2.RECEIPT_LNO" & vbCrLf _
            & "     and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
            & "     and APTINVH1.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
            & "     and APTINVH1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & "   group by APTINVH5.RECEIPT_NO, APTINVH5.RECEIPT_LNO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTVDEF2 & vbCrLf _
            & "    Set QTY_INV = R1.QTY_INV, AMT_INV = R1.AMT_INV" & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"

        ''Dim SPECIAL_FIX As Boolean = (ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso Format(Now, "yyyyMMdd") = "20260210")
        'Dim SPECIAL_FIX As Boolean = (Format(Now, "yyyyMMdd") = "20260210")
        'If SPECIAL_FIX Then
        '    ASCMAIN1.sql = "" _
        '        & "Begin" & vbCrLf _
        '        & " Declare Cursor C1 is" & vbCrLf _
        '        & "  Select APTINVH5.RECEIPT_NO, APTINVH5.RECEIPT_LNO" & vbCrLf _
        '        & " , Sum ((NVL(APTINVH5.INV_QTY,0) - NVL(WJZINVS.INV_QTY,0))) QTY_INV" & vbCrLf _
        '        & " , Sum (NVL(APTINVH5.INV_QTY,0) * NVL(APTINVH5.INV_COST,0)) AMT_INV" & vbCrLf _
        '        & "   from APTINVH5,APTINVH1," & ICTVDEF2 & " ICTVDEF2, (SELECT * FROM WJZINVS WHERE NVL(WJZINVS.RECS,0) >= 2 AND NVL(WJZINVS.PCT,0) >= 100) WJZINVS" & vbCrLf _
        '        & "   where APTINVH5.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
        '        & "     and APTINVH5.RECEIPT_NO = ICTVDEF2.RECEIPT_NO" & vbCrLf _
        '        & "     and APTINVH5.RECEIPT_LNO = ICTVDEF2.RECEIPT_LNO" & vbCrLf _
        '        & "     and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
        '        & "     and APTINVH1.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
        '        & "     and APTINVH1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        '        & "     and WJZINVS.VNOMAX (+) = APTINVH5.VOUCHER_NO AND WJZINVS.RECEIPT_NO (+) = APTINVH5.RECEIPT_NO AND WJZINVS.RECEIPT_LNO (+) = APTINVH5.RECEIPT_LNO" & vbCrLf _
        '        & "   group by APTINVH5.RECEIPT_NO, APTINVH5.RECEIPT_LNO;" & vbCrLf _
        '        & " Begin" & vbCrLf _
        '        & "  For R1 in C1 Loop" & vbCrLf _
        '        & "   Update " & ICTVDEF2 & vbCrLf _
        '        & "    Set QTY_INV = R1.QTY_INV, AMT_INV = R1.AMT_INV" & vbCrLf _
        '        & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
        '        & "  End Loop;" & vbCrLf _
        '        & " End;" & vbCrLf _
        '        & "End;"
        'End If



        ASCDATA1.ExecuteSQL()

        ' DO WE NEED TO SPLIT F VS T, AND DO WE NEED WHERE CLAUSES LOOKING FOR F OR T?

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO" & vbCrLf _
            & " , Sum (CASE WHEN APTACRC1.ACCRUAL_CODE = 'FRT' THEN NVL(APTINVH7.TOTAL_INV,0) ELSE 0 END) AMT_INV_LANDG" & vbCrLf _
            & " , Sum (CASE WHEN APTACRC1.ACCRUAL_CODE = 'TRF' THEN NVL(APTINVH7.TOTAL_INV,0) ELSE 0 END) AMT_INV_TOOLG" & vbCrLf _
            & "   from APTINVH7,APTACRC1,APTINVH1," & ICTVDEF2 & " ICTVDEF2" & vbCrLf _
            & "   where APTINVH7.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
            & "     and APTINVH7.CTL_NO = APTACRC1.CTL_NO" & vbCrLf _
            & "     and APTACRC1.RECEIPT_NO = ICTVDEF2.RECEIPT_NO" & vbCrLf _
            & "     and APTACRC1.RECEIPT_LNO = ICTVDEF2.RECEIPT_LNO" & vbCrLf _
            & "     and APTACRC1.ACCRUAL_CODE IN ('FRT','TRF')" & vbCrLf _
            & "     and APTINVH1.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
            & "     and APTINVH1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & "     and APTINVH1.INV_STATUS <> 'D'" & vbCrLf _
            & "   group by APTACRC1.RECEIPT_NO, APTACRC1.RECEIPT_LNO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTVDEF2 & vbCrLf _
            & "    Set AMT_INV_LANDG = R1.AMT_INV_LANDG, AMT_INV_TOOLG = R1.AMT_INV_TOOLG" & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO and RECEIPT_LNO = R1.RECEIPT_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Alter Table " & ICTVDEF2 & " Add VOUCHER_NO VARCHAR2(10)")
        ASCDATA1.ExecuteSQL("Update " & ICTVDEF2 & " X Set VOUCHER_NO = (Select Min (APTINVH5.VOUCHER_NO) from APTINVH5,APTINVH1 where APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO and APTINVH1.OPS_YYYYPP >= '" & FYP & "' and APTINVH1.OPS_YYYYPP <= '" & RYP & "' and APTINVH5.RECEIPT_NO = X.RECEIPT_NO and APTINVH5.RECEIPT_LNO = X.RECEIPT_LNO)")

        ASCMAIN1.sql = "Select * from " & ICTVDEF1 & " ICTVDEF1"
        Create_TDA(dst.Tables.Add, "ICTVDEF1", "**", 0, False, "", 1)

        'ASCMAIN1.sql = "Select * from " & ICTVDEF2 & " ICTVDEF2"

        ASCMAIN1.sql = "Select ICTVDEF2.*" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_MAKE_BUY STD_MB" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_VCOST STD_VCOST" & vbCrLf _
            & ", NVL(ICTCOSTA.ITEM_COST_MATLS,0) + NVL(ICTCOSTA.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA.ITEM_COST_OVRHDI,0) STD_MATLS" & vbCrLf _
            & ", ICTCOSTA.ITEM_COST_LANDG STD_LANDG, ICTCOSTA.ITEM_COST_TOOLG STD_TOOLG, ICTCOSTA.ITEM_COST_OVRHD STD_OVRHD" & vbCrLf _
            & ", ICTCOSTA_REC.ITEM_COST_MAKE_BUY REC_MB" & vbCrLf _
            & ", ICTCOSTA_REC.ITEM_COST_VCOST REC_VCOST" & vbCrLf _
            & ", NVL(ICTCOSTA_REC.ITEM_COST_MATLS,0) + NVL(ICTCOSTA_REC.ITEM_COST_LANDGI,0) + NVL(ICTCOSTA_REC.ITEM_COST_TOOLGI,0) + NVL(ICTCOSTA_REC.ITEM_COST_OVRHDI,0) REC_MATLS" & vbCrLf _
            & ", ICTCOSTA_REC.ITEM_COST_LANDG REC_LANDG, ICTCOSTA_REC.ITEM_COST_TOOLG REC_TOOLG, ICTCOSTA_REC.ITEM_COST_OVRHD REC_OVRHD" & vbCrLf _
            & $" from {ICTVDEF2} ICTVDEF2, ICTCOSTA, ICTCOSTA ICTCOSTA_REC" & vbCrLf _
            & "where ICTCOSTA.ITEM_CODE = ICTVDEF2.ITEM_CODE" & vbCrLf _
            & $"  And ICTCOSTA.OPS_YYYYPP = '{RYP}'" & vbCrLf _
            & "  and ICTCOSTA_REC.ITEM_CODE = ICTVDEF2.ITEM_CODE" & vbCrLf _
            & "  and ICTCOSTA_REC.OPS_YYYYPP = ICTVDEF2.OPS_YYYYPP"
        Create_TDA(dst.Tables.Add, "ICTVDEF2", "**", 0, False, "", 0)
        ' THESE NEXT 2 FIELDS OCCUR IN THE DATA AT ODG
        With dst.Tables("ICTVDEF2").Columns
            .Add("AMT_REC", GetType(System.Decimal), "ISNULL(QTY_REC,0) * ISNULL(PO_COST,0)")
            .Add("AMT_ACCRUAL_OPEN", GetType(System.Decimal), "(ISNULL(QTY_REC,0) - ISNULL(QTY_INV,0)) * ISNULL(PO_COST,0)")

            'NOT USED, AND IF IT WERE USED, IT SHOULD BE A REFLECTION OF TOTAL STD
            '.Add("EXT_USED", GetType(System.Decimal), "IIF(ISNULL(VOUCHER_NO,'X')='X',ISNULL(AMT_REC,0),ISNULL(AMT_INV,0)) * ISNULL(QTY_USED,0) / ISNULL(QTY_REC,0)")

            .Add("EXT_USED_PCT", GetType(System.Decimal), "ISNULL(QTY_USED,0) / ISNULL(QTY_REC,0)")
            .Add("EXT_USED_VCOST", GetType(System.Decimal), "EXT_USED_PCT * (ISNULL(AMT_INV,0)+ISNULL(AMT_ACCRUAL_OPEN,0))")
            .Add("EXT_USED_MATLS", GetType(System.Decimal), "EXT_USED_PCT * ISNULL(EXT_COST_MATLS,0)")
            '.Add("EXT_USED_LANDG", GetType(System.Decimal), "EXT_USED_PCT * IIF(CTL_STATUS_LANDG = '1', ISNULL(AMT_INV_LANDG,0),ISNULL(AMT_ACC_LANDG,0))") 
            ' DOES NOT WORK ANY MORE EVER SINCE WE WENT TO SPLIT ACCRUALS
            .Add("EXT_USED_LANDG", GetType(System.Decimal), "EXT_USED_PCT * AMT_USED_LANDG")
            .Add("EXT_USED_TOOLG", GetType(System.Decimal), "EXT_USED_PCT * AMT_USED_TOOLG")

            '.Add("EXT_USED", GetType(System.Decimal), "EXT_USED_VCOST + EXT_USED_MATLS + EXT_USED_LANDG")
            .Add("EXT_USED", GetType(System.Decimal), "EXT_USED_VCOST + EXT_USED_MATLS + EXT_USED_LANDG + EXT_USED_TOOLG")

            .Add("INV_COST", GetType(System.Decimal), "IIF(ISNULL(QTY_INV,0) = 0, 0, ISNULL(AMT_INV,0)/ISNULL(QTY_INV,0))")
        End With
        dst.Tables("ICTVDEF2").Columns("EXT_USED_PCT").DefaultValue = 0

        Create_Relation("ICTVDEF1", "ICTVDEF2", "ITEM_CODE")
        With dst.Tables("ICTVDEF1").Columns
            .Add("QTY_USED", GetType(System.Int32), "SUM(CHILD(ICTVDEF1_ICTVDEF2).QTY_USED)")
            .Add("EXT_USED_VCOST", GetType(System.Decimal), "SUM(CHILD(ICTVDEF1_ICTVDEF2).EXT_USED_VCOST)")
            .Add("EXT_USED_MATLS", GetType(System.Decimal), "SUM(CHILD(ICTVDEF1_ICTVDEF2).EXT_USED_MATLS)")
            .Add("EXT_USED_LANDG", GetType(System.Decimal), "SUM(CHILD(ICTVDEF1_ICTVDEF2).EXT_USED_LANDG)")
            .Add("EXT_USED_TOOLG", GetType(System.Decimal), "SUM(CHILD(ICTVDEF1_ICTVDEF2).EXT_USED_TOOLG)")
            .Add("EXT_USED", GetType(System.Decimal), "SUM(CHILD(ICTVDEF1_ICTVDEF2).EXT_USED)")
            .Add("VAR_VCOST", GetType(System.Decimal), "ISNULL(EXT_USED_VCOST,0) - ISNULL(EXT_STD_VCOST,0)")
            .Add("VAR_MATLS", GetType(System.Decimal), "ISNULL(EXT_USED_MATLS,0) - ISNULL(EXT_STD_MATLS,0)")
            .Add("VAR_LANDG", GetType(System.Decimal), "ISNULL(EXT_USED_LANDG,0) - ISNULL(EXT_STD_LANDG,0)")
            .Add("VAR_TOOLG", GetType(System.Decimal), "ISNULL(EXT_USED_TOOLG,0) - ISNULL(EXT_STD_TOOLG,0)")
            .Add("VARIANCE", GetType(System.Decimal), "VAR_VCOST + VAR_MATLS + VAR_LANDG + VAR_TOOLG")
        End With

        Create_Relation("ICTCOST1", "ICTVDEF1", "COST_CATGY_CODE")
        With dst.Tables("ICTCOST1").Columns
            .Add("VARIANCE", GetType(System.Decimal), "SUM(CHILD(ICTCOST1_ICTVDEF1).VARIANCE)")
            .Add("VAR_VCOST", GetType(System.Decimal), "SUM(CHILD(ICTCOST1_ICTVDEF1).VAR_VCOST)")
            .Add("VAR_MATLS", GetType(System.Decimal), "SUM(CHILD(ICTCOST1_ICTVDEF1).VAR_MATLS)")
            .Add("VAR_LANDG", GetType(System.Decimal), "SUM(CHILD(ICTCOST1_ICTVDEF1).VAR_LANDG)")
            .Add("VAR_TOOLG", GetType(System.Decimal), "SUM(CHILD(ICTCOST1_ICTVDEF1).VAR_TOOLG)")
        End With

        With dst.Tables.Add("ICTCOST2")
            .Columns.Add("COST_CATGY_CODE")
            .Columns.Add("COLLECTION_CODE")
            .PrimaryKey = New DataColumn() {.Columns("COST_CATGY_CODE"), .Columns("COLLECTION_CODE")}
        End With

        Create_Relation("ICTCOST2", "ICTVDEF1", "COST_CATGY_CODE,COLLECTION_CODE")
        With dst.Tables("ICTCOST2").Columns
            .Add("VARIANCE", GetType(System.Decimal), "SUM(CHILD(ICTCOST2_ICTVDEF1).VARIANCE)")
            .Add("VAR_VCOST", GetType(System.Decimal), "SUM(CHILD(ICTCOST2_ICTVDEF1).VAR_VCOST)")
            .Add("VAR_MATLS", GetType(System.Decimal), "SUM(CHILD(ICTCOST2_ICTVDEF1).VAR_MATLS)")
            .Add("VAR_LANDG", GetType(System.Decimal), "SUM(CHILD(ICTCOST2_ICTVDEF1).VAR_LANDG)")
            .Add("VAR_TOOLG", GetType(System.Decimal), "SUM(CHILD(ICTCOST2_ICTVDEF1).VAR_TOOLG)")
        End With

        With dst.Tables.Add("ICTVDEF3")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ITEM_DESC")
            .Columns.Add("COST_CATGY_CODE")
            .Columns.Add("ITEM_COST_TOTAL", GetType(System.Decimal))
            .Columns.Add("QTY_USED", GetType(System.Int64))
            .Columns("QTY_USED").DefaultValue = 0

            For I As Integer = 1 To 6
                .Columns.Add("AGE" & CStr(I), GetType(System.Decimal))
                .Columns("AGE" & CStr(I)).DefaultValue = 0
            Next

            .Columns.Add("AGEX", GetType(System.Decimal), "AGE1+AGE2+AGE3+AGE4+AGE5+AGE6")
            .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
        End With

        grdICTVDEF3.DataSource = dst.Tables("ICTVDEF3")

        If grdICTVDEF3.DisplayLayout.Bands(0).Summaries.Count = 0 Then
            Create_Summary(grdICTVDEF3, New String() {"QTY_USED", "AGE1", "AGE2", "AGE3", "AGE4", "AGE5", "AGE6", "AGEX"})
        End If

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        Create_TDA(dst.Tables.Add, "ICTCOLL1", "*", 0, False)
        Fill_Records("ICTCOLL1")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        Fill_Records("ICTVDEF1")

        ASCMAIN1.sql = "" _
            & "BEGIN " _
            & " DECLARE  " _
            & "  CURSOR C1 IS SELECT * FROM " & ICTVDEF1 & " WHERE QTY_ONH > 0;" _
            & "  BAL NUMBER;" _
            & " BEGIN " _
            & "  FOR R1 IN C1 LOOP" _
            & "   BAL := R1.QTY_ONH;" _
            & "   BEGIN " _
            & "    DECLARE " _
            & "     CURSOR C2 IS SELECT * FROM " & ICTVDEF2 _
            & "                  WHERE ITEM_CODE = R1.ITEM_CODE " _
            & "                  ORDER BY RECEIPT_DATE DESC, SOURCE_DOC_NO DESC, RECEIPT_NO DESC FOR UPDATE;" _
            & "     QTY NUMBER;" _
            & "    BEGIN " _
            & "     FOR R2 IN C2 LOOP" _
            & "      IF R2.QTY_REC > BAL THEN" _
            & "       QTY := BAL;" _
            & "       BAL := 0;" _
            & "      ELSE" _
            & "       QTY := R2.QTY_REC;" _
            & "       BAL := BAL - R2.QTY_REC;" _
            & "      END IF;" _
            & "      UPDATE " & ICTVDEF2 & " SET QTY_USED = QTY WHERE CURRENT OF C2;" _
            & "      IF BAL <= 0 THEN" _
            & "       EXIT;" _
            & "      END IF;" _
            & "     END LOOP; " _
            & "    END; " _
            & "   END;" _
            & "  END LOOP; " _
            & " END; " _
            & "END;"
        ASCDATA1.ExecuteSQL()

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        ASCMAIN1.sql = "Update " & ICTVDEF2 & " Set AGE_DAYS = TO_DATE('" & Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy") & "') - RECEIPT_DATE"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from " & ICTVDEF2 & " where NVL(QTY_USED,0) = 0"

        Fill_Records("ICTVDEF2")

        For Each rowICTVDEF1 As DataRow In dst.Tables("ICTVDEF1").Rows
            Dim ITEM_CODE As String = rowICTVDEF1.Item("ITEM_CODE")
            'If ASCMAIN1.Running_in_VS Then
            '    If ITEM_CODE = "CH012V21USA" Then Stop
            'End If
            Dim ITEM_COST_TOTAL As Decimal = Val(rowICTVDEF1.Item("ITEM_COST_TOTAL") & "")
            Dim ITEM_COST_VCOST As Decimal = Val(rowICTVDEF1.Item("ITEM_COST_VCOST") & "")
            Dim ITEM_COST_MATLS As Decimal = Val(rowICTVDEF1.Item("ITEM_COST_MATLS") & "") ' note that this field actually has the implicit amounts added in
            Dim QTY_ONH As Int32 = Val(rowICTVDEF1.Item("QTY_ONH") & "")
            Dim QTY_USED As Int32 = Val(rowICTVDEF1.Item("QTY_USED") & "")

            Dim BAL As Int32 = QTY_ONH - QTY_USED

            If BAL <> 0 Then
                Dim rowICTVDEF2 As DataRow = dst.Tables("ICTVDEF2").NewRow
                With rowICTVDEF2
                    .Item("QTY_REC") = BAL
                    .Item("QTY_USED") = BAL
                    .Item("PO_ORDER_NO") = "Forced"
                    .Item("RECEIPT_NO") = "000000"
                    .Item("RECEIPT_LNO") = 0
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("ITEM_COST_STD") = ITEM_COST_TOTAL
                    .Item("PO_COST") = ITEM_COST_VCOST
                    .Item("EXT_COST_MATLS") = BAL * ITEM_COST_MATLS

                    Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", ITEM_CODE)
                    If rowICTCOSTC IsNot Nothing Then
                        .Item("STD_MB") = rowICTCOSTC.Item("ITEM_COST_MAKE_BUY")
                        .Item("STD_VCOST") = Val(rowICTCOSTC.Item("ITEM_COST_VCOST") & "")
                        .Item("STD_MATLS") = Val(rowICTCOSTC.Item("ITEM_COST_MATLS") & "") + Val(rowICTCOSTC.Item("ITEM_COST_LANDGI") & "") + Val(rowICTCOSTC.Item("ITEM_COST_TOOLGI") & "") + Val(rowICTCOSTC.Item("ITEM_COST_OVRHDI") & "")
                        .Item("STD_LANDG") = Val(rowICTCOSTC.Item("ITEM_COST_LANDG") & "")
                        .Item("STD_TOOLG") = Val(rowICTCOSTC.Item("ITEM_COST_TOOLG") & "")
                        .Item("STD_OVRHD") = Val(rowICTCOSTC.Item("ITEM_COST_OVRHD") & "")

                        .Item("COST_CATGY_CODE") = rowICTCOSTC.Item("COST_CATGY_CODE")

                    End If
                End With

                dst.Tables("ICTVDEF2").Rows.Add(rowICTVDEF2)
            End If

            If Val(rowICTVDEF1.Item("VARIANCE") & "") <> 0 Then
                Dim rowICTIVAR1 As DataRow = dst.Tables("ICTIVAR1").NewRow
                rowICTIVAR1.Item("ITEM_CODE") = ITEM_CODE
                rowICTIVAR1.Item("OPS_YYYYPP") = RYP
                rowICTIVAR1.Item("PV_DEF") = Val(rowICTVDEF1.Item("VAR_VCOST") & "")
                rowICTIVAR1.Item("MV_DEF") = Val(rowICTVDEF1.Item("VAR_MATLS") & "")
                rowICTIVAR1.Item("FV_DEF") = Val(rowICTVDEF1.Item("VAR_LANDG") & "")
                rowICTIVAR1.Item("TV_DEF") = Val(rowICTVDEF1.Item("VAR_TOOLG") & "")
                dst.Tables("ICTIVAR1").Rows.Add(rowICTIVAR1)
            End If
        Next

        ASCDATA1.Keep_Rows(dst.Tables("ICTVDEF2"), "QTY_USED <> 0")

        dst.Tables("ICTVDEF3").Rows.Clear()
        For Each rowICTVDEF2 As DataRow In dst.Tables("ICTVDEF2").Select("")
            Dim ITEM_CODE As String = rowICTVDEF2.Item("ITEM_CODE")
            ' If ITEM_CODE = "200608" Then Stop
            Dim rowICTVDEF3 As DataRow = dst.Tables("ICTVDEF3").Rows.Find(New String() {ITEM_CODE})
            If rowICTVDEF3 Is Nothing Then
                rowICTVDEF3 = dst.Tables("ICTVDEF3").NewRow
                rowICTVDEF3.Item("ITEM_CODE") = ITEM_CODE
                rowICTVDEF3.Item("ITEM_DESC") = rowICTVDEF2.GetParentRow("ICTVDEF1_ICTVDEF2").Item("ITEM_DESC")
                rowICTVDEF3.Item("COST_CATGY_CODE") = rowICTVDEF2.GetParentRow("ICTVDEF1_ICTVDEF2").Item("COST_CATGY_CODE")
                rowICTVDEF3.Item("ITEM_COST_TOTAL") = rowICTVDEF2.GetParentRow("ICTVDEF1_ICTVDEF2").Item("ITEM_COST_TOTAL")
                dst.Tables("ICTVDEF3").Rows.Add(rowICTVDEF3)
            End If

            Dim QTY_USED As Int32 = Val(rowICTVDEF2.Item("QTY_USED") & "")
            Dim ITEM_COST_TOTAL As Decimal = Val(rowICTVDEF3.Item("ITEM_COST_TOTAL") & "")
            Dim AGE_DAYS As Int32 = Val(rowICTVDEF2.Item("AGE_DAYS") & "")
            Dim AGEX As String = ""
            If AGE_DAYS <= 60 Then
                AGEX = "AGE1"
            ElseIf AGE_DAYS <= 120 Then
                AGEX = "AGE2"
            ElseIf AGE_DAYS <= 180 Then
                AGEX = "AGE3"
            ElseIf AGE_DAYS <= 360 Then
                AGEX = "AGE4"
            ElseIf AGE_DAYS <= 720 Then
                AGEX = "AGE5"
            Else
                AGEX = "AGE6"
            End If
            rowICTVDEF3.Item("QTY_USED") = Val(rowICTVDEF3.Item("QTY_USED")) + QTY_USED
            rowICTVDEF3.Item(AGEX) = Val(rowICTVDEF3.Item(AGEX)) + QTY_USED * ITEM_COST_TOTAL ' + Val(rowICTVDEF2.Item("EXT_USED"))
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTVDEF1", New String() {"COST_CATGY_CODE", "COLLECTION_CODE"}).Select("")
            Dim COST_CATGY_CODE As String = row.Item("COST_CATGY_CODE")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
            dst.Tables("ICTCOST2").Rows.Add(New String() {COST_CATGY_CODE, COLLECTION_CODE})
        Next

        Prepare_GL_Interface("ICVD")

        EnforceConstraints(True)
    End Sub

    Public Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Decimal
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))

        For Each rowICTCOST2 As DataRow In dst.Tables("ICTCOST2").Select _
            ("VAR_VCOST <> 0 OR VAR_MATLS <> 0 OR VAR_LANDG <> 0 OR VAR_TOOLG <> 0", "")

            Dim COST_CATGY_CODE As String = rowICTCOST2.Item("COST_CATGY_CODE")
            Dim rowICTCOST1 As DataRow = dst.Tables("ICTCOST1").Rows.Find(COST_CATGY_CODE)

            Dim COLLECTION_CODE As String = rowICTCOST2.Item("COLLECTION_CODE")
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

            For Each VAR As String In New String() {"PPV", "MUV", "FPV", "TPV"}
                If VAR = "PPV" Then
                    DETL_POSTING_AMT = Val(rowICTCOST2.Item("VAR_VCOST") & "")
                ElseIf VAR = "MUV" Then
                    DETL_POSTING_AMT = Val(rowICTCOST2.Item("VAR_MATLS") & "")
                ElseIf VAR = "FPV" Then
                    DETL_POSTING_AMT = Val(rowICTCOST2.Item("VAR_LANDG") & "")
                ElseIf VAR = "TPV" Then
                    DETL_POSTING_AMT = Val(rowICTCOST2.Item("VAR_TOOLG") & "")
                Else
                    DETL_POSTING_AMT = 0
                End If

                DETL_POSTING_AMT = Round(DETL_POSTING_AMT, 2)

                If DETL_POSTING_AMT <> 0 Then
                    For Each YP As String In New String() _
                        {RYP, ASCMAIN1.Period_Calc(RYP, 1)}

                        Dim ACCT_CODE As String = ""
                        For I As Int32 = 0 To 1
                            If I = 0 Then
                                ACCT_CODE = rowICTCOST1.Item("ACCT_CODE_" & VAR & "_DEF") & ""
                            Else
                                ACCT_CODE = rowICTCOST1.Item("ACCT_CODE_" & VAR & "_EXP") & ""
                            End If

                            If I = 1 Then
                                DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
                            End If

                            Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
                            rowGLTINTF1("OPS_YYYYPP") = YP
                            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                            JOURNAL_LNO += 1
                            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                            rowGLTINTF1("ACCT_CODE") = ACCT_CODE
                            Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            Dim SEG3_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            Dim SEG4_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

                            If ROWs("ICTPARM1").Item("IC_PARM_EXP_SEG4") & "" = "1" Then
                                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                                    SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                                Else
                                    SEG4_CODE = COLLECTION_CODE
                                End If
                            End If

                            rowGLTINTF1("SEG2_CODE") = SEG2_CODE
                            rowGLTINTF1("SEG3_CODE") = SEG3_CODE
                            rowGLTINTF1("SEG4_CODE") = SEG4_CODE

                            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
                            rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                            rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
                            rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
                            rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
                            rowGLTINTF1("DETL_CVX_NO") = COST_CATGY_CODE & ":" & COLLECTION_CODE
                            rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
                            rowGLTINTF1("DETL_CVX_REF_NO") = VAR
                            rowGLTINTF1("DETL_DESC") = DBNull.Value
                            rowGLTINTF1("DETL_CVX_TYPE") = DBNull.Value
                            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                            ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                        Next
                    Next
                End If


            Next
        Next

        Return JOURNAL_NO

    End Function


    Overrides Sub Verify_Special_Pre(ByVal eItemKey As String)
        If eItemKey = "Proceed" Or eItemKey = "Update" Then
            Dim RYP As String = Absx1.cmbFor("RYP").Value
            RYP = Mid(RYP, 1, 4) & Mid(RYP, 6, 2)
            'ASCMAIN1.sql = $"Select * from GLTJRNL1 where OPS_YYYYPP =  and JOURNAL_TYPE = 'ICVD'"
            ASCMAIN1.sql = $"Select Distinct GLTDETL1.JOURNAL_NO, GLTDETL1.OPS_YYYYPP from GLTDETL1,GLTJRNL1 where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO and GLTJRNL1.JOURNAL_TYPE= 'ICVD' and GLTDETL1.OPS_YYYYPP = '{RYP}'"
            Dim rows() As DataRow = ASCDATA1.GetDataTable().Select("")
            If rows.Length > 1 Then
                If MsgBox($"It appears that a Reversal and a Book J/E" & vbCrLf & $" has already been posted to { Absx1.cmbFor("RYP").Value}" & vbCrLf & vbCrLf & "Continue Anyway?", MsgBoxStyle.YesNo, "Verification") = vbNo Then
                    EMsg &= "Please check the Period"
                End If
            End If
        End If
    End Sub

    Function CopyDataToSSG(TABLE_NAME As String, Optional columnCountToKeep As Int32 = -1, Optional sort_by As String = "") As String

        ' Declare SSG Objects

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        ' Parameters

        'Dim Start_Row As Integer = 5

        ' Save Workbook as FILENAME
        Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & XNO & ".xlsx"

        ASCMAIN1.Progress("Now Creating Intermediary SSG Workbook")

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        oSheet = oWB.Worksheets.Add
        oSheet.Name = "Data"
        Dim dvw As DataView = dst.Tables(TABLE_NAME).DefaultView
        If sort_by <> "" Then
            dvw.Sort = sort_by
        End If

        'oSheet.Range(0, 0).CopyFromDataTable(dst.Tables(TABLE_NAME), SpreadsheetGear.Data.SetDataFlags.None) ' headings included
        oSheet.Range(0, 0).CopyFromDataTable(dvw.ToTable, SpreadsheetGear.Data.SetDataFlags.None) ' headings included
        'oSheet.Range(0, 0).CopyFromDataTable(dvw.ToTable, SpreadsheetGear.Data.SetDataFlags.AllText) ' headings included

        If columnCountToKeep <> -1 Then
            oSheet.Range(0, columnCountToKeep, 0, oSheet.UsedRange.ColumnCount - 1).EntireColumn.Delete()
        End If

        oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        oWB.Close()

        oWB = Nothing

        Return XLS_FILENAME
    End Function
End Class