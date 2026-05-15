Public Class SORPOLOG
    Dim SOTPOLOG As String = ""
    Dim tblSOTORDR0 As DataTable

    Private Sub SORPOLOG_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        If grdSOTORDR0.DisplayLayout.Bands(0).Summaries.Count = 0 Then
            Create_Summary(grdSOTORDR0, "SEL")
            Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})
            Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        End If

    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ASCMAIN1.sql = "Select * from SOTORDR0 where ROWNUM < 1"
        tblSOTORDR0 = ASCDATA1.GetDataTable
        tblSOTORDR0.Columns.Add("SEL")
        tblSOTORDR0.Columns("SEL").DefaultValue = "0"
        tblSOTORDR0.TableName = "SOTORDR0"

        grdSOTORDR0.DataSource = tblSOTORDR0
        With grdSOTORDR0.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("ORDR_GROUP_NO").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.Gold
                End With

                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = System.Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        'Create_Summary(grdSOTORDR0, "SEL")
        'Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})
        'Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        grdSOTORDR0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Show_Filter(grdSOTORDR0, True)
        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = False

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Absx1.cmbFor("RYP").Visible = (optSELECT.Value = "P")

        Setup_SOTORDR0()
    End Sub

    Protected Overrides Sub Build_Workfile()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        If optSELECT.Value = "P" Then
            '  Stop ' CHECK RYP
        Else
            RYP = ASCMAIN1.CYP
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

    Private Sub optSELECT_ValueChanged(sender As Object, e As EventArgs) Handles optSELECT.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        'Set_Read_Only_for_ctl(Absx1.cmbFor("RYP"), Not (optSELECT.Value = "P"))

        Absx1.cmbFor("RYP").Visible = (optSELECT.Value = "P")
        If optSELECT.Value = "P" Then
        Else
            For Each GROW As UltraWinGrid.UltraGridRow In Absx1.cmbFor("RYP").Rows
                If GROW.Cells("OPS_YYYYPP").Value = ASCMAIN1.CYP Then
                    Absx1.cmbFor("RYP").SelectedRow = GROW
                End If
            Next
        End If


        Setup_SOTORDR0()

    End Sub

    Sub Setup_SOTORDR0()

        If optSELECT.Value = "P" Then
            Dim LEGEND As String = Absx1.cmbFor("RYP").Value & ""
            Dim RYP As String = Mid(LEGEND, 1, 4) & Mid(LEGEND, 6, 2)
            ASCMAIN1.sql = "Select * from SOTORDR0" & vbCrLf _
                & " where OPS_YYYYPP_MIN >= '" & RYP & "' and OPS_YYYYPP_MAX <= '" & RYP & "'"
        Else
            ASCMAIN1.sql = "Select * from SOTORDR0" & vbCrLf _
                & " where OPS_YYYYPP_MIN >= '" & ASCMAIN1.CYP & "' and OPS_YYYYPP_MAX <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & " union " & vbCrLf _
                & " Select * from SOTORDR0 where ORDR_STATUS in ('O','P')"

        End If

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        tblSOTORDR0.Rows.Clear()
        tblSOTORDR0.Merge(tbl)

    End Sub

    Sub Create_Pivot()

        Dim SQLW As String = ""
        SQLW &= SQLA_filter("CUST_CODE", "SOTORDR1")
        SQLW &= SQLA_filter("SREP_CODE", "SOTORDR1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "SOTORDR1")

        If tblSOTORDR0.Select("SEL = '1'").Length <> 0 Then
            Dim ORDR_GROUP_NOs As String = ""
            For Each row As DataRow In tblSOTORDR0.Select("SEL = '1'")
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                ORDR_GROUP_NOs &= ",'" & ORDR_GROUP_NO & "'"
            Next
            SQLW &= " and SOTORDR1.ORDR_GROUP_NO in (" & Mid(ORDR_GROUP_NOs, 2) & ")"
        End If

        ASCMAIN1.Progress("Now Creating Workbook")

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

        ws = wb.Worksheets("Data")


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  Trunc (SYSDATE) REPORT_DATE" & vbCrLf _
            & ", TO_CHAR(SYSDATE,'HH:MI:SS') REPORT_TIME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE COMPANY" & vbCrLf _
            & ", SOTINVH1.ORDR_CUST_PO CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO ORDER_NUM" & vbCrLf _
            & ", SOTINVH1.ORDR_TYPE_CODE ORDR_TYPE" & vbCrLf _
            & ", SOTTYPE1.ORDR_TYPE_DESC ORDR_TYPE_DESC" & vbCrLf _
            & ", DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','BASIC','P','PROMO','?') ORDER_CLASS" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE ORDER_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE REQSHIP_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_CANCEL_DATE CANCEL_DATE" & vbCrLf _
            & ", DECODE(SOTINVH1.INV_TYPE,'C','Returned','I','Shipped','?') ORDER_STATUS" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY ORDER_QTY" & vbCrLf _
            & ", ARTCUST1.CUST_NAME CHAIN" & vbCrLf _
            & ", SOTINVH1.CUST_CODE CUS1" & vbCrLf _
            & ", SOTINVH1.CUST_STORE_NO CS2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
            & ", 'INTERPARFUM' FRANCHISE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE BRAND_NUM" & vbCrLf _
            & ", ICTBRAN1.BRAND_NAME BRAND_NAME" _
            & ", ICTCOLL1.COLLECTION_NAME LINE" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE ITEM" & vbCrLf _
            & ", ICTITEM1.ITEM_EAN_CODE EAN_UPC" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT ALT_ITEM_NUM" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.PROD_CODE PROD_CLASS" & vbCrLf _
            & ", ICTPROD1.PROD_DESC PROD_CLASS_DESC" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE SPDESC" & vbCrLf _
            & ", SOTINVH2.ITEM_RETAIL_PRICE MSRP" & vbCrLf _
            & ", SOTINVH2.ITEM_RETAIL_PRICE * SOTINVH2.ORDR_QTY_SHIP VAL_SRP" & vbCrLf _
            & ", SOTINVH2.ORDR_UNIT_PRICE * SOTINVH2.ORDR_QTY_SHIP VAL_NET" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP SHIP_QTY" & vbCrLf _
            & ", SOTINVH1.INV_DATE SHIPPED_DATE" & vbCrLf _
            & ", SOTINVH1.INV_NO INVOICE_NUM" & vbCrLf _
            & ", SOTSREP1.SREP_NAME KAM_NAME" & vbCrLf _
            & ", CASE WHEN SOTORDR1.ORDR_SHIP_DATE IS NULL THEN 'RETURN' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') = '" & RYP & "' THEN 'CURRENT' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') < '" & RYP & "' THEN 'CRYOVR' ELSE 'FUTURE' END END END DATE_RANGE" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE INDC_DATE" & vbCrLf _
            & ", SOTREAS1.REASON_DESC RETURN_REASON_DESCRIPTION" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " from SOTINVH1,SOTINVH2,SOTORDR1,SOTORDR2,ICTITEM1,ICTCOLL1,ICTBRAN1,ICTPROD1,SOTSREP1,SOTREAS1,SOTTYPE1,ARTCUST1,ARTCUST2" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTINVH2.INV_LNO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO (+) = SOTINVH1.CUST_STORE_NO" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE (+) = SOTINVH1.SREP_CODE" & vbCrLf _
            & "   and SOTREAS1.REASON_CODE (+) = SOTINVH1.REASON_CODE" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
            & Replace(Replace(SQLW, "SOTORDR1.", "SOTINVH1."), "SOTINVH1.ORDR_GROUP_NO", "SOTORDR1.ORDR_GROUP_NO")


        ASCMAIN1.sql &= "UNION Select " & vbCrLf _
            & "  Trunc (SYSDATE) REPORT_DATE" & vbCrLf _
            & ", TO_CHAR(SYSDATE,'HH:MI:SS') REPORT_TIME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE COMPANY" & vbCrLf _
            & ", SOTINVH1.ORDR_CUST_PO CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO ORDER_NUM" & vbCrLf _
            & ", SOTINVH1.ORDR_TYPE_CODE ORDR_TYPE" & vbCrLf _
            & ", SOTTYPE1.ORDR_TYPE_DESC ORDR_TYPE_DESC" & vbCrLf _
            & ", DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','BASIC','P','PROMO','?') ORDER_CLASS" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE ORDER_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE REQSHIP_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_CANCEL_DATE CANCEL_DATE" & vbCrLf _
            & ", DECODE(SOTINVH1.INV_TYPE,'C','Returned','I','Shipped','?') ORDER_STATUS" & vbCrLf _
            & ", 0 ORDER_QTY" & vbCrLf _
            & ", ARTCUST1.CUST_NAME CHAIN" & vbCrLf _
            & ", SOTINVH1.CUST_CODE CUS1" & vbCrLf _
            & ", SOTINVH1.CUST_STORE_NO CS2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
            & ", 'INTERPARFUM' FRANCHISE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE BRAND_NUM" & vbCrLf _
            & ", ICTBRAN1.BRAND_NAME BRAND_NAME" _
            & ", ICTCOLL1.COLLECTION_NAME LINE" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE ITEM" & vbCrLf _
            & ", ICTITEM1.ITEM_EAN_CODE EAN_UPC" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT ALT_ITEM_NUM" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.PROD_CODE PROD_CLASS" & vbCrLf _
            & ", ICTPROD1.PROD_DESC PROD_CLASS_DESC" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE SPDESC" & vbCrLf _
            & ", SOTINVH2.ITEM_RETAIL_PRICE MSRP" & vbCrLf _
            & ", SOTINVH2.ITEM_RETAIL_PRICE * SOTINVH2.ORDR_QTY_SHIP VAL_SRP" & vbCrLf _
            & ", SOTINVH2.ORDR_UNIT_PRICE * SOTINVH2.ORDR_QTY_SHIP VAL_NET" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP SHIP_QTY" & vbCrLf _
            & ", SOTINVH1.INV_DATE SHIPPED_DATE" & vbCrLf _
            & ", SOTINVH1.INV_NO INVOICE_NUM" & vbCrLf _
            & ", SOTSREP1.SREP_NAME KAM_NAME" & vbCrLf _
            & ", CASE WHEN SOTORDR1.ORDR_SHIP_DATE IS NULL THEN 'RETURN' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') = '" & RYP & "' THEN 'CURRENT' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') < '" & RYP & "' THEN 'CRYOVR' ELSE 'FUTURE' END END END DATE_RANGE" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE INDC_DATE" & vbCrLf _
            & ", SOTREAS1.REASON_DESC RETURN_REASON_DESCRIPTION" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " from SOTINVH1,SOTINVH2,SOTORDR1,ICTITEM1,ICTCOLL1,ICTBRAN1,ICTPROD1,SOTSREP1,SOTREAS1,SOTTYPE1,ARTCUST1,ARTCUST2" & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = 'C'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO (+) = SOTINVH1.CUST_STORE_NO" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE (+) = SOTINVH1.SREP_CODE" & vbCrLf _
            & "   and SOTREAS1.REASON_CODE (+) = SOTINVH1.REASON_CODE" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
            & Replace(Replace(SQLW, "SOTORDR1.", "SOTINVH1."), "SOTINVH1.ORDR_GROUP_NO", "SOTORDR1.ORDR_GROUP_NO")

        ASCMAIN1.sql &= " UNION Select " & vbCrLf _
            & "  Trunc (SYSDATE) REPORT_DATE" & vbCrLf _
            & ", TO_CHAR(SYSDATE,'HH:MI:SS') REPORT_TIME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE COMPANY" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO ORDER_NUM" & vbCrLf _
            & ", SOTORDR1.ORDR_TYPE_CODE ORDR_TYPE" & vbCrLf _
            & ", SOTTYPE1.ORDR_TYPE_DESC ORDR_TYPE_DESC" & vbCrLf _
            & ", DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','BASIC','P','PROMO','?') ORDER_CLASS" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE ORDER_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE REQSHIP_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_CANCEL_DATE CANCEL_DATE" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_STATUS,'O','Open','P','In Pick','C','Cancelled','?') ORDER_STATUS" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,'P',SOTORDR2.ORDR_QTY_PICK,'C',SOTORDR2.ORDR_QTY_CANC,SOTORDR2.ORDR_QTY) ORDER_QTY" & vbCrLf _
            & ", SOTORDR1.CUST_NAME CHAIN" & vbCrLf _
            & ", SOTORDR1.CUST_CODE CUS1" & vbCrLf _
            & ", SOTORDR1.CUST_STORE_NO CS2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
            & ", 'INTERPARFUM' FRANCHISE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE BRAND_NUM" & vbCrLf _
            & ", ICTBRAN1.BRAND_NAME BRAND_NAME" _
            & ", ICTCOLL1.COLLECTION_NAME LINE" & vbCrLf _
            & ", SOTORDR2.ITEM_CODE ITEM" & vbCrLf _
            & ", ICTITEM1.ITEM_EAN_CODE EAN_UPC" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT ALT_ITEM_NUM" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.PROD_CODE PROD_CLASS" & vbCrLf _
            & ", ICTPROD1.PROD_DESC PROD_CLASS_DESC" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE SPDESC" & vbCrLf _
            & ", SOTORDR2.ITEM_RETAIL_PRICE MSRP" & vbCrLf _
            & ", SOTORDR2.ITEM_RETAIL_PRICE * DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,'P',SOTORDR2.ORDR_QTY_PICK,'C',SOTORDR2.ORDR_QTY_CANC,SOTORDR2.ORDR_QTY) VAL_SRP" & vbCrLf _
            & ", SOTORDR2.ORDR_UNIT_PRICE * DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY_OPEN,'P',SOTORDR2.ORDR_QTY_PICK,'C',SOTORDR2.ORDR_QTY_CANC,SOTORDR2.ORDR_QTY) VAL_NET" & vbCrLf _
            & ", NULL SHIP_QTY" & vbCrLf _
            & ", NULL SHIPPED_DATE" & vbCrLf _
            & ", NULL INVOICE_NUM" & vbCrLf _
            & ", SOTSREP1.SREP_NAME KAM_NAME" & vbCrLf _
            & ", CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') = '" & RYP & "' THEN 'CURRENT' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') < '" & RYP & "' THEN 'CRYOVR' ELSE 'FUTURE' END END DATE_RANGE" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE INDC_DATE" & vbCrLf _
            & ", SOTREAS1.REASON_DESC RETURN_REASON_DESCRIPTION" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,ICTITEM1,ICTCOLL1,ICTBRAN1,ICTPROD1,SOTSREP1,SOTREAS1,SOTTYPE1,ARTCUST1,ARTCUST2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE (+) = SOTORDR1.SREP_CODE" & vbCrLf _
            & "   and SOTREAS1.REASON_CODE (+) = SOTORDR1.REASON_CODE" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and (SOTORDR1.ORDR_STATUS IN ('O','P') or (SOTORDR1.ORDR_STATUS = 'C' and SOTORDR1.ORDR_YYYYPP_CLOSED = '" & RYP & "'))" & vbCrLf _
            & SQLW
        '  & "   and SOTORDR1.CUST_CODE IN ('SAKS','BELK','BONTON','BLOOMIES')"

        ASCMAIN1.sql &= " UNION Select " & vbCrLf _
            & "  Trunc (SYSDATE) REPORT_DATE" & vbCrLf _
            & ", TO_CHAR(SYSDATE,'HH:MI:SS') REPORT_TIME" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE COMPANY" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO CUST_PO" & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO ORDER_NUM" & vbCrLf _
            & ", SOTORDR1.ORDR_TYPE_CODE ORDR_TYPE" & vbCrLf _
            & ", SOTTYPE1.ORDR_TYPE_DESC ORDR_TYPE_DESC" & vbCrLf _
            & ", DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','BASIC','P','PROMO','?') ORDER_CLASS" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE ORDER_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_DATE REQSHIP_DATE" & vbCrLf _
            & ", SOTORDR1.ORDR_CANCEL_DATE CANCEL_DATE" & vbCrLf _
            & ", DECODE(SOTORDR2.ORDR_RELEASE,'D','Deleted','S','Suppress','R','Re-Order','Cancelled') ORDER_STATUS" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_CANC ORDER_QTY" & vbCrLf _
            & ", SOTORDR1.CUST_NAME CHAIN" & vbCrLf _
            & ", SOTORDR1.CUST_CODE CUS1" & vbCrLf _
            & ", SOTORDR1.CUST_STORE_NO CS2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME DOORNAME" & vbCrLf _
            & ", 'INTERPARFUM' FRANCHISE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE BRAND_NUM" & vbCrLf _
            & ", ICTBRAN1.BRAND_NAME BRAND_NAME" _
            & ", ICTCOLL1.COLLECTION_NAME LINE" & vbCrLf _
            & ", SOTORDR2.ITEM_CODE ITEM" & vbCrLf _
            & ", ICTITEM1.ITEM_EAN_CODE EAN_UPC" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT ALT_ITEM_NUM" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.PROD_CODE PROD_CLASS" & vbCrLf _
            & ", ICTPROD1.PROD_DESC PROD_CLASS_DESC" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE SPDESC" & vbCrLf _
            & ", SOTORDR2.ITEM_RETAIL_PRICE MSRP" & vbCrLf _
            & ", SOTORDR2.ITEM_RETAIL_PRICE * SOTORDR2.ORDR_QTY_CANC VAL_SRP" & vbCrLf _
            & ", SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY_CANC VAL_NET" & vbCrLf _
            & ", NULL SHIP_QTY" & vbCrLf _
            & ", NULL SHIPPED_DATE" & vbCrLf _
            & ", NULL INVOICE_NUM" & vbCrLf _
            & ", SOTSREP1.SREP_NAME KAM_NAME" & vbCrLf _
            & ", CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') = '" & RYP & "' THEN 'CURRENT' ELSE CASE WHEN TO_CHAR(SOTORDR1.ORDR_SHIP_DATE,'YYYYMM') < '" & RYP & "' THEN 'CRYOVR' ELSE 'FUTURE' END END DATE_RANGE" & vbCrLf _
            & ", SOTORDR1.ORDR_ARRIVAL_DATE INDC_DATE" & vbCrLf _
            & ", SOTREAS1.REASON_DESC RETURN_REASON_DESCRIPTION" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,ICTITEM1,ICTCOLL1,ICTBRAN1,ICTPROD1,SOTSREP1,SOTREAS1,SOTTYPE1,ARTCUST1,ARTCUST2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE (+) = SOTORDR1.SREP_CODE" & vbCrLf _
            & "   and SOTREAS1.REASON_CODE (+) = SOTORDR1.REASON_CODE" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and (SOTORDR1.ORDR_STATUS IN ('O','P')  or (SOTORDR1.ORDR_STATUS = 'F' and SOTORDR1.ORDR_YYYYPP_UPDATED = '" & RYP & "'))" & vbCrLf _
            & "   and SOTORDR2.ORDR_QTY_CANC <> 0" & vbCrLf _
            & SQLW
        '  & "   and SOTORDR1.CUST_CODE IN ('SAKS','BELK','BONTON','BLOOMIES')"

        DataTable = ASCDATA1.GetDataTable

        r = 0
        For Each row As DataRow In DataTable.Select("")
            r += 1
            ws.Range("A" & CStr(3 + r) & ":AM" & CStr(3 + r)).Value2 = row.ItemArray
        Next
        wb.Names.Add("DataPivotBase", "=DATA!$A$3:$AM$" & CStr(3 + DataTable.Rows.Count))

        'xlSourceRange = ws.Range("A4:Q4")
        'xlDestRange = ws.Range("A4:Q" & CStr(3 + DataTable.Rows.Count))
        'xlSourceRange.Copy(xlDestRange)
        'ws.Cells(1, 3).Value = Now

        '   excel.Run("ResetData")

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "PO_Log"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsm"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
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