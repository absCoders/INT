Public Class SOROPEN1
    Private ASTSRPT1bu As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Range_Events(grpORDR_DATE_BOOKED)
        Range_Events(grpORDR_SHIP_DATE)
        Range_Events(grpORDR_CANCEL_DATE)
        Range_Events(grpORDR_DATE_SHIPPED)



        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        If Absx1.optFor("OPTEDI").Value = "N" Then
            sql_filter &= " and NVL(SOTORDR1.ORDR_SOURCE,'?') <> 'E'"
        ElseIf Absx1.optFor("OPTEDI").Value <> "A" Then
            sql_filter &= " and NVL(SOTORDR1.ORDR_SOURCE,'?') = '" & Absx1.optFor("OPTEDI").Value & "'"
        End If

        If Absx1.optFor("OPTSTATUS").Value = "A" Then
            sql_filter &= " and SOTORDR1.ORDR_STATUS <> 'D'"
        Else
            If Absx1.optFor("OPTSTATUS").Value = "X" Then
                sql_filter &= " and SOTORDR1.ORDR_STATUS in ('O','P')"
            Else
                sql_filter &= " and SOTORDR1.ORDR_STATUS = '" & Absx1.optFor("OPTSTATUS").Value & "'"
            End If
        End If

        Page0.Add("Status: " & Absx1.optFor("OPTSTATUS").Text)
        Page0.Add("Source: " & Absx1.optFor("OPTEDI").Text)

        'If Absx1.optFor("OPTASN").Value = "S" Then sql_filter &= " and ICTSTYL1.CUST_CODE is Null"
        'If Absx1.optFor("OPTASN").Value = "N" Then sql_filter &= " and ICTSTYL1.CUST_CODE is Not Null"

        Dim OPT_G As String = "SOTORDR1.ORDR_GROUP_NO"
        Dim OPT_G1 As String = "SOTORDR1.ORDR_CUST_PO"
        Dim OPT_G2 As String = "SOTORDR1.CUST_STORE_NO"
        Dim OPT_G3 As String = "SOTORDR1.WHSE_CODE"
        Dim OPT_G4 As String = "SOTORDR1.ORDR_DATE"
        Dim OPT_G5 As String = "SOTORDR1.ORDR_SHIP_DATE"
        Dim OPT_G6 As String = "SOTORDR1.ORDR_CANCEL_DATE"

        ' 
        Dim OPT_S As String = "SOTORDR2.ITEM_CODE"

        If Absx1.optFor("OPTDTL").Value = "3" Then
            OPT_G = "'0000000000'"
        End If

        If Absx1.optFor("OPTDTL").Value <> "2" Then
            OPT_S = "'X'"
        End If

        Dim SHOW_AMT As String = "NVL(SOTORDR2.ORDR_UNIT_PRICE,0)"
        If Absx1.optFor("OPTSHOWAMT").Value = "R" Then
            Page0.Add("Showing Amount as Extended Retail Price")
            SHOW_AMT = "NVL(ICTITEM1.ITEM_RETAIL_PRICE,0)"
        ElseIf Absx1.optFor("OPTSHOWAMT").Value = "C" Then
            Page0.Add("Showing Amount as Extended Std Cost")
            SHOW_AMT = "NVL(ICTITEM1.ITEM_COST_STD,0)"
        Else
            Page0.Add("Showing Amount as Extended Net Price")
        End If


        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", " & OPT_G & " ORDR_GROUP_NO" & vbCrLf _
            & ", " & OPT_G1 & " ORDR_CUST_PO" & vbCrLf _
            & ", " & OPT_G2 & " CUST_STORE_NO" & vbCrLf _
            & ", " & OPT_G3 & " WHSE_CODE" & vbCrLf _
            & ", " & OPT_G4 & " ORDR_DATE" & vbCrLf _
            & ", " & OPT_G5 & " ORDR_SHIP_DATE" & vbCrLf _
            & ", " & OPT_G6 & " ORDR_CANCEL_DATE" & vbCrLf _
            & ", " & OPT_S & " ITEM_CODE" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY,0)) ORDR_QTY" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0)) ORDR_QTY_ALLO" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY,0) * " & SHOW_AMT & ") ORDR_AMT" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0) * " & SHOW_AMT & ") ORDR_AMT_ALLO" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0) * " & SHOW_AMT & ") ORDR_AMT_OPEN" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0) * " & SHOW_AMT & ") ORDR_AMT_PICK" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0) * " & SHOW_AMT & ") ORDR_AMT_SHIP" & vbCrLf _
            & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * " & SHOW_AMT & ") ORDR_AMT_CANC" & vbCrLf _
            & ", 0.01 ORDR_UNIT_PRICE" & vbCrLf _
            & " from " & "SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & Get_Dates("O")) & vbCrLf _
            & " group by " & IIf(sql_GROUP_BY_cols = "", "'" & "O" & "'", sql_GROUP_BY_cols) & vbCrLf _
            & ", " & OPT_G & vbCrLf _
            & ", " & OPT_G1 & vbCrLf _
            & ", " & OPT_G2 & vbCrLf _
            & ", " & OPT_G3 & vbCrLf _
            & ", " & OPT_G4 & vbCrLf _
            & ", " & OPT_G5 & vbCrLf _
            & ", " & OPT_G6 & vbCrLf _
            & ", " & OPT_S & vbCrLf


        'sql = "Select " & sql_SELECT_cols & vbCrLf _
        '    & ", " & OPT_G & " ORDR_GROUP_NO" & vbCrLf _
        '    & ", " & OPT_S & " ITEM_CODE" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY,0)) ORDR_QTY" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0)) ORDR_QTY_ALLO" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY,0) * " & SHOW_AMT & ") ORDR_AMT" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0) * " & SHOW_AMT & ") ORDR_AMT_ALLO" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0) * " & SHOW_AMT & ") ORDR_AMT_OPEN" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0) * " & SHOW_AMT & ") ORDR_AMT_PICK" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0) * " & SHOW_AMT & ") ORDR_AMT_SHIP" & vbCrLf _
        '    & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * " & SHOW_AMT & ") ORDR_AMT_CANC" & vbCrLf _
        '    & ", 0.01 ORDR_UNIT_PRICE" & vbCrLf _
        '    & " from " & "SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
        '    & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & Get_Dates("O")) & vbCrLf _
        '    & " group by " & IIf(sql_GROUP_BY_cols = "", "'" & "O" & "'", sql_GROUP_BY_cols) & vbCrLf _
        '    & ", " & OPT_G & vbCrLf _
        '    & ", " & OPT_S & vbCrLf

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")


        ' ASTSRPT1bu = ASCMAIN1.Temp_Table(sql)





        sql = "Select ORDR_GROUP_NO, ITEM_CODE" & ASTSRPT1_sum_columns _
            & " from " & ASTSRPT1 & " group by ORDR_GROUP_NO, ITEM_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRY", 4))

        sql = "Select X.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & " from (Select Distinct ITEM_CODE from " & ASTSRPT1 & ") X, ICTITEM1" _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTITEMX", 1))

        If Absx1.optFor("OPTDTL").Value = "2" Then
            ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " Set ITEM_CODE = NULL")
        End If

        sql = "" _
            & "Select ORDR_GROUP_NO, CUST_CODE, ORDR_DATE" & vbCrLf _
            & ", ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_CUST_PO, ORDR_DEPT" & vbCrLf _
            & ", ORDR_CNT, ORDR_CNT_OPEN, ORDR_CNT_PICK, CUST_DC_NO" & vbCrLf _
            & ", SALES_DIVISION_CODE, WHSE_CODE from SOTORDR0 " & vbCrLf _
            & " where ORDR_GROUP_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRX", 1))

    End Sub

    Function Get_Dates(TYPE As String) As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_DATE_SHIPPED"}
            Dim grp As Infragistics.Win.Misc.UltraGroupBox = DirectCast(Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent, Infragistics.Win.Misc.UltraGroupBox)
            Dim R As String = Replace(grp.Text, " Range", "")
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                Dim z As String = Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy")
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " >= '" & z & "'"
                Page0.Add(R & " from " & z)
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                Dim z As String = Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy")
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " <= '" & z & "'"
                Page0.Add(R & "to " & z)
            End If
        Next
        sql = Replace(sql, "SOTORDR1.ORDR_DATE_BOOKED", "TRUNC(SOTORDR1.INIT_DATE)")
        Return sql
    End Function

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = 0")
        ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = TRUNC(100 * ORDR_AMT / ORDR_QTY) / 100 where ORDR_QTY <> 0")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""

        Page0.Add("Report Detail: " & Absx1.optFor("OPTDTL").Text)

        Page0.Add("Status: " & Absx1.optFor("OPTSTATUS").Text)
        If Absx1.optFor("OPTSTATUS").Value <> "A" Then SUBT &= Absx1.optFor("OPTSTATUS").Text & ", "

        Page0.Add("Orders: " & Absx1.optFor("OPTEDI").Text)
        If Absx1.optFor("OPTEDI").Value <> "A" Then SUBT &= Absx1.optFor("OPTEDI").Text & ", "

        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_DATE_SHIPPED"}
            Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                Z &= " from First"
            Else
                Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
            End If
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                Z &= " to Last"
            Else
                Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
            End If
            Page0.Add(Z)
        Next

        ' Stop ' CR_Rpt.ParameterFields("LVLS").SetCurrentValue(CStr(cfmax + 1))
        CR_params.Add("CHKDTL", IIf(Absx1.optFor("OPTDTL").Value = "2", "1", "0"))
        CR_params.Add("OPTDTL", Absx1.optFor("OPTDTL").Value)
        CR_params.Add("OPTSHOWAMT", Absx1.optFor("OPTSHOWAMT").Value)

        If Absx1.optFor("OPTSTATUS").Value = "X" Then
            RPT_TITLE = "Sales Order Report"
        End If
        Generate_Report(RPT, RPT_TITLE, SUBT)

        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        'Create_Relation("ARTCUSTX", "ARTATBR1", "CUST_CODE")
        'With dst.Tables("ARTATBR1").Columns
        '    .Add("CUST_NAME", GetType(System.String), "PARENT.CUST_NAME")
        '    .Add("TRADE_CLASS_CODE", GetType(System.String), "PARENT.TRADE_CLASS_CODE")
        '    .Add("TERM_DESC", GetType(System.String), "PARENT.TERM_DESC")
        'End With

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")

        'grdASTEXPT1.DataSource = dst.Tables("SOTORDRX")
        grdASTEXPT1.Text = "Open Orders"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        'Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer")
        'Set_DX_Column(grdASTEXPT1, "CUST_NAME", "Customer Name", 30)
        Set_DX_Column(grdASTEXPT1, "ORDR_GROUP_NO", "Group", 110)

        Set_DX_Column(grdASTEXPT1, "ORDR_CUST_PO", "Cust PO", 110)
        'Set_DX_Column(grdASTEXPT1, "CUST_DC_NO", "DC", 80)
        Set_DX_Column(grdASTEXPT1, "WHSE_CODE", "Whse", 60)
        Set_DX_Column(grdASTEXPT1, "ORDR_DATE", "Order Dt", 90, "MM/dd/yy")
        Set_DX_Column(grdASTEXPT1, "ORDR_SHIP_DATE", "Ship Dt", 90, "MM/dd/yy")
        Set_DX_Column(grdASTEXPT1, "ORDR_CANCEL_DATE", "Cancel Dt", 90, "MM/dd/yy")
        'Set_DX_Column(grdASTEXPT1, "INV_DUE_DATE", "Due Date")
        'Set_DX_Column(grdASTEXPT1, "INV_TOTAL_AMOUNT", "Orig Amt", 120, "#,##0.00", "Sum")
        'Set_DX_Column(grdASTEXPT1, "INV_BALANCE", "Balance Due", 120, "#,##0.00", "Sum")



        'If Absx1.optFor("OPTID").Value = "I" Then
        '    Set_DX_Column(grdASTEXPT1, "AGE", "Days", 50, "#,##0")
        'Else
        '    Set_DX_Column(grdASTEXPT1, "DUE", "Days", 50, "#,##0")
        'End If
        'Set_DX_Column(grdASTEXPT1, "INV_BALANCE_1", Absx1.txtFor("LBL_DAYS1").Text, 120, "#,##0.00", "Sum")
        'Set_DX_Column(grdASTEXPT1, "INV_BALANCE_2", Absx1.txtFor("LBL_DAYS2").Text, 120, "#,##0.00", "Sum")
        'Set_DX_Column(grdASTEXPT1, "INV_BALANCE_3", Absx1.txtFor("LBL_DAYS3").Text, 120, "#,##0.00", "Sum")
        'Set_DX_Column(grdASTEXPT1, "INV_BALANCE_4", Absx1.txtFor("LBL_DAYS4").Text, 120, "#,##0.00", "Sum")

        Dim WorkbookView1 As SpreadsheetGear.Windows.Forms.WorkbookView

        Dim tabpage_exists = False
        If tabDataExports.Tabs.Count > 1 AndAlso tabDataExports.Tabs(1).Key = "Open Orders SSG" Then
            tabpage_exists = True
        End If
        If tabpage_exists Then 'tabDataExports.Tabs.Contains("Open Orders SSG") Then
            WorkbookView1 = DirectCast(tabDataExports.Tabs("Open Orders SSG").TabPage.Controls(0), SpreadsheetGear.Windows.Forms.WorkbookView)

        Else

            tabDataExports.Tabs.Add("Open Orders SSG")
            tabDataExports.Tabs("Open Orders SSG").Text = tabDataExports.Tabs("Open Orders SSG").Key
            WorkbookView1 = New SpreadsheetGear.Windows.Forms.WorkbookView
            WorkbookView1.Name = "WorkbookView1"
            WorkbookView1.Parent = tabDataExports.Tabs("Open Orders SSG").TabPage
            WorkbookView1.Dock = DockStyle.Fill

            Dim btn As New Infragistics.Win.Misc.UltraButton
            btn.Parent = WorkbookView1 ' tabDataExports.Tabs("Open Orders SSG").TabPage
            btn.Text = ""
            btn.Appearance.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "32\" & "Excel" & ".png")
            btn.Visible = True
            btn.Top = 0
            btn.Left = 0
            btn.Width = 25
            btn.Height = 25
            AddHandler btn.Click, AddressOf btnExcel_Click
        End If


        WorkbookView1.GetLock()
        WorkbookView1.ActiveWorksheet.Cells.Clear()
        Load_DataTable_into_SGXLS(1, 1, dst.Tables("ASTSRPT1"), WorkbookView1.ActiveWorksheet, grdASTEXPT1, Nothing, "", "")
        'Load_DataTable_into_SGXLS(1, 1, dst.Tables("SOTORDRX"), WorkbookView1.ActiveWorksheet, grdASTEXPT1, Nothing, "", "")
        WorkbookView1.ReleaseLock()


    End Sub


    Private Sub btnExcel_Click(sender As Object, e As EventArgs)
        Dim WorkbookView1 As SpreadsheetGear.Windows.Forms.WorkbookView
        WorkbookView1 = DirectCast(tabDataExports.Tabs("Open Orders SSG").TabPage.Controls(0), SpreadsheetGear.Windows.Forms.WorkbookView)
        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & ASCMAIN1.Next_Control_No($"{Me.Name}.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("ITEM_CODE")

                If Absx1.optFor("OPTDTL").Value = "2" And Val(rowASTDSQLA("SEQUENCE") & "") <> 0 Then
                    EMsg &= "You Must NOT Sort by Style when showing Details"
                End If

                EMsg &= TAC.TACMAIN1.Check_Permissions(Me) ' for FS
        End Select
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("SREP_CODE")
                ' rowASTDSQLA.Item("CODE_VALUES") = TAC.TACMAIN1.SREP_CODE
                rowASTDSQLA.Item("CODE_VALUES") = Join(TAC.TACMAIN1.SREP_CODEs.ToArray, ",")
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Sub Pivot_Prepare_PreProcess(dt As DataTable)
        dt.Columns.Remove("ORDR_GROUP_NO")
        dt.Columns.Remove("ITEM_CODE")
    End Sub

    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows

            For J As Integer = 1 To SEQs
                Dim COLNAME As String = "G" & Format(J)
                rowASTSRPT1.Item(COLNAME) = Split(rowASTSRPT1.Item(COLNAME), ":")(1)
            Next
        Next



        ' MAYBE FIX ASTRPT1 for second option

        ''If Absx1.optFor("OPTDTL").Value = "2" Then

        ''    dst.Tables("ASTSRPT1").Rows.Clear()
        ''    ASCMAIN1.sql = "Select * from " & ASTSRPT1bu
        ''    Fill_Records("ASTSRPT1", "", True, ASCMAIN1.sql)
        ''End If
        ' grdASTSRPT1

        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Hidden = True

        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Header.Caption = "Group No"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_CUST_PO").Header.Caption = "Customer PO"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.Caption = "DC"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("WHSE_CODE").Header.Caption = "Whse"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_DATE").Header.Caption = "Order Dt"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_SHIP_DATE").Header.Caption = "Ship Dt"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_CANCEL_DATE").Header.Caption = "Cancel DT"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY").Header.Caption = "Ord Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO").Header.Caption = "Allo Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY_OPEN").Header.Caption = "Open Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY_PICK").Header.Caption = "Pick Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").Header.Caption = "Ship Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_QTY_CANC").Header.Caption = "Canc Units"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT").Header.Caption = "Order Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT_ALLO").Header.Caption = "Allo Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT_OPEN").Header.Caption = "Open Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT_PICK").Header.Caption = "Pick Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT_SHIP").Header.Caption = "Ship Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_AMT_CANC").Header.Caption = "Canc Amount"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Header.Caption = "Ordr Unit Price"

        'ORDR_UNIT_PRICE
    End Sub
End Class