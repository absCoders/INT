Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.forms
Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles

Public Class ICFINVA1

    Dim RYP As String
    Dim RDT As Date
    Dim FYP As String
    Dim FDT As Date
    Dim ICTINVA1 As String = ""
    Dim IOT As New Dictionary(Of String, String)
    Dim sqlw As String = ""
    Dim sql_ICTITEMX As String = ""
    Dim MOSMAX As Integer = 0
    Private printDialog1 As System.Windows.Forms.PrintDialog

    Private viewICTINVAT As DataView

    Private selectedPriceCatgyCode As String = String.Empty
    Private selectedProdCatgyCode As String = String.Empty
    Private selectedVendor As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select MOS from ICTINVA1"
            MyBase.Create_TDA(.Tables.Add, "ICTINVA0", _
            CommonConstants.CustomSqlIndicator, 0, False, String.Empty, 0)
            With .Tables("ICTINVA0").Columns
                .Add("I_ALL", GetType(System.Int32))
                .Add("I_STK", GetType(System.Int32))
                .Add("I_RXP", GetType(System.Int32))
                .Add("I_FRZ", GetType(System.Int32))
                .Add("S_ALL", GetType(System.Double))
                .Add("S_STK", GetType(System.Double))
                .Add("S_RXP", GetType(System.Double))
                .Add("S_FRZ", GetType(System.Double))
                .Add("L_ALL", GetType(System.Int32))
                .Add("L_STK", GetType(System.Int32))
                .Add("L_RXP", GetType(System.Int32))
                .Add("L_FRZ", GetType(System.Int32))
                .Add("Q_ALL", GetType(System.Int32))
                .Add("Q_STK", GetType(System.Int32))
                .Add("Q_RXP", GetType(System.Int32))
                .Add("Q_FRZ", GetType(System.Int32))
            End With

            Create_TDA(.Tables.Add, "ICTINVA1", "*")
            With .Tables("ICTINVA1").Columns
                .Add("ITEM_UOM", GetType(System.String))
                .Add("PO_QTY_MULT", GetType(System.Int32))
                .Add("PRICE_CATGY_COST_TOTAL", GetType(System.Decimal))
                .Add("ITEM_UPC_CODE", GetType(System.String))
                .Add("ITEM_OPC_CODE", GetType(System.String))

                ' Item Attributes
                .Add("ITEM_BASE_CURVE", GetType(System.Decimal))
                .Add("ITEM_DIAMETER", GetType(System.Decimal))
                .Add("ITEM_SPHERE_POWER", GetType(System.Decimal))
                .Add("ITEM_CYLINDER", GetType(System.Decimal))
                .Add("ITEM_AXIS", GetType(System.Int32))
                .Add("ITEM_COLOR", GetType(System.String))
                .Add("ITEM_ADD_POWER", GetType(System.String))
                .Add("ITEM_ADD_DOM_NON", GetType(System.String))
                .Add("ITEM_LEFT_RIGHT", GetType(System.String))
                .Add("ITEM_CENTER_THICKNESS", GetType(System.Decimal))
                .Add("ITEM_BLANK_SIZE", GetType(System.Decimal))
                .Add("ITEM_INDEX_REF", GetType(System.Decimal))
                .Add("ITEM_OZD", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" _
            & ", POTORDR1.VEND_CODE, POTORDR1.VEND_NAME, POTORDR1.PO_DATE_ORDERED" _
            & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_COST" _
            & " from POTORDR1,POTORDR2 " _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" _
            & "   and POTORDR2.PO_STATUS = 'O'" _
            & "   and POTORDR2.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" _
            & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_DATE" _
            & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_UNIT_PRICE" _
            & " from SOTORDR1,SOTORDR2 " _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & "   and SOTORDR1.ORDR_STATUS = 'O'" _
            & "   and SOTORDR2.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" _
            & ", SOTINVH2.INV_DATE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME" _
            & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
            & ", SOTINVH2.ORDR_UNIT_PRICE_PATIENT " _
            & " from SOTINVH2,SOTINVH1,ARTCUST1 " _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE " _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" _
            & "   and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 3)
            .Tables("SOTINVHX").Columns.Add("ORDR_LINE_AMT", GetType(System.Double))
            .Tables("SOTINVHX").Columns("ORDR_LINE_AMT").Expression = "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)"

            ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME" _
                & " from APTVEND1 "
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)
            .Tables("APTVEND1").Columns.Add("ITEMS", GetType(System.Int32))

            ASCMAIN1.sql = "Select ICTPCAT1.PRICE_CATGY_CODE, ICTPCAT1.PRICE_CATGY_DESC" _
                & ", ICTPCAT1.VEND_CODE" _
                & " from ICTPCAT1 "
            Create_TDA(.Tables.Add, "ICTPCAT1", "**", 0, False, "", 1)
            .Tables("ICTPCAT1").Columns.Add("ITEMS", GetType(System.Int32))

            ASCMAIN1.sql = "Select ICTCATG1.PROD_CATGY_CODE, ICTCATG1.PROD_CATGY_DESC" _
            & " from ICTCATG1 "
            Create_TDA(.Tables.Add, "ICTCATG1", "**", 0, False, "", 1)
            .Tables("ICTCATG1").Columns.Add("ITEMS", GetType(System.Int32))

            ASCMAIN1.sql = "Select VEND_CODE CODE_VALUE, VEND_NAME DESC_VALUE" _
            & " from APTVEND1 "
            Create_TDA(.Tables.Add, "ICTINVAT", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC" _
            & " from ICTCATL1 where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "V", 1)
        End With

        grdICTINVA0.DataSource = dst.Tables("ICTINVA0")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        grdICTINVA1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide

        grdAPTVEND1.DataSource = dst.Tables("APTVEND1")
        grdICTCATG1.DataSource = dst.Tables("ICTCATG1")
        grdICTPCAT1.DataSource = dst.Tables("ICTPCAT1")

        grdICTINVA1.DataSource = dst.Tables("ICTINVA1")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdICTINVAT.DataSource = dst.Tables("ICTINVAT")

        viewICTINVAT = New DataView(dst.Tables("ICTINVAT"))

        With grdICTINVA0.DisplayLayout.Bands(0)
            .Columns("S_ALL").Format = "#,##0"
            .Columns("S_STK").Format = "#,##0"
            .Columns("S_RXP").Format = "#,##0"
            .Columns("S_FRZ").Format = "#,##0"
        End With

        'Call Get_PARM("GLTPARM1")

        'grdGLTSUMJ1.DisplayLayout.Bands("GLTSUMJ2").SummaryFooterCaption = "Totals for Journal Type: [JOURNAL_TYPE] [JOURNAL_TYPE_DESC]"

        'Create_Lookup("GLTACCT1")

        Call Create_Summary(grdICTINVA0, "I_ALL")
        Call Create_Summary(grdICTINVA0, "I_STK")
        Call Create_Summary(grdICTINVA0, "I_RXP")
        Call Create_Summary(grdICTINVA0, "I_FRZ")

        Call Create_Summary(grdICTINVA0, "S_ALL")
        Call Create_Summary(grdICTINVA0, "S_STK")
        Call Create_Summary(grdICTINVA0, "S_RXP")
        Call Create_Summary(grdICTINVA0, "S_FRZ")

        Call Create_Summary(grdICTINVA0, "L_ALL")
        Call Create_Summary(grdICTINVA0, "L_STK")
        Call Create_Summary(grdICTINVA0, "L_RXP")
        Call Create_Summary(grdICTINVA0, "L_FRZ")

        Call Create_Summary(grdICTINVA0, "Q_ALL")
        Call Create_Summary(grdICTINVA0, "Q_STK")
        Call Create_Summary(grdICTINVA0, "Q_RXP")
        Call Create_Summary(grdICTINVA0, "Q_FRZ")

        Call Create_Summary(grdICTINVA1, "ITEM_CODE", "Count")
        Call Create_Summary(grdICTINVA1, "QOH")
        Call Create_Summary(grdICTINVA1, "QPO")
        Call Create_Summary(grdICTINVA1, "QSO")
        Call Create_Summary(grdICTINVA1, "S00")
        Call Create_Summary(grdICTINVA1, "L00")
        Call Create_Summary(grdICTINVA1, "Q00")

        Call Create_Summary(grdICTINVAT, "CODE_VALUE", "Count")

        Call Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Call Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Call Create_Summary(grdSOTINVHX, "ORDR_LINE_AMT")

        Call Create_Summary(grdPOTORDRX, "PO_ORDER_LNO", "Count")
        Call Create_Summary(grdPOTORDRX, "PO_QTY_OPN")

        Call Create_Summary(grdSOTORDRX, "ORDR_LNO", "Count")
        Call Create_Summary(grdSOTORDRX, "ORDR_QTY_OPEN")

        Call Create_Summary(grdAPTVEND1, "ITEMS")
        Call Create_Summary(grdAPTVEND1, "VEND_CODE", "Count")
        Call Create_Summary(grdICTCATG1, "ITEMS")
        Call Create_Summary(grdICTCATG1, "PROD_CATGY_CODE", "Count")
        Call Create_Summary(grdICTPCAT1, "ITEMS")
        Call Create_Summary(grdICTPCAT1, "PRICE_CATGY_CODE", "Count")

        grdICTINVA0.DisplayLayout.UseFixedHeaders = True
        With grdICTINVA0.DisplayLayout.Bands("ICTINVA0")
            '.Columns("ITEM_CODE").Header.Fixed = True
            .Groups(0).Header.Fixed = True
        End With

        grdICTINVA1.DisplayLayout.UseFixedHeaders = True
        With grdICTINVA1.DisplayLayout.Bands("ICTINVA1")
            '.Columns("ITEM_CODE").Header.Fixed = True
            .Groups("Item").Header.Fixed = True
        End With

        With grdICTINVA0.DisplayLayout
            .Bands("ICTINVA0").SortedColumns.Clear()
            .Bands("ICTINVA0").SortedColumns.Add("MOS", False)
        End With

        With grdICTCATG1.DisplayLayout
            .Bands("ICTCATG1").SortedColumns.Clear()
            .Bands("ICTCATG1").SortedColumns.Add("PROD_CATGY_CODE", False)
        End With

        With grdAPTVEND1.DisplayLayout
            .Bands("APTVEND1").SortedColumns.Clear()
            .Bands("APTVEND1").SortedColumns.Add("VEND_CODE", False)
        End With

        With grdICTPCAT1.DisplayLayout
            .Bands("ICTPCAT1").SortedColumns.Clear()
            .Bands("ICTPCAT1").SortedColumns.Add("PRICE_CATGY_CODE", False)
        End With

        chtICTINVAT.Axis.X.ScrollScale.Visible = True
        chtICTINVAT.Axis.Y.ScrollScale.Visible = True

        chtICTINVAT.Axis.X.ScrollScale.Scale = 1 ' 0.25
        chtICTINVAT.Axis.Y.ScrollScale.Scale = 1 ' 0.25
        Me.trkbrXAxis.Value = chtICTINVAT.Axis.X.ScrollScale.Scale * 100
        Me.trkbrYAxis.Value = chtICTINVAT.Axis.Y.ScrollScale.Scale * 100
        chtICTINVAT.EnableCrossHair = True


        Dim modelStyle As String() = System.Enum.GetNames(GetType(ColorModels))
        Dim s As String
        For Each s In modelStyle
            Me.comboBox1.Items.Add(s)
        Next s

        Me.comboBox1.SelectedItem = Me.comboBox1.Items(Me.comboBox1.FindString(System.Enum.GetName(GetType(ColorModels), chtICTINVAT.ColorModel.ModelStyle), 0))

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        dteEnding.Value = Now.Date
        Do While CDate(dteEnding.Value).DayOfWeek <> DayOfWeek.Friday
            dteEnding.Value = CDate(dteEnding.Value).AddDays(-1)
        Loop

        IOT.Add("N", "STK")
        IOT.Add("P", "RXP")
        IOT.Add("X", "FRZ")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If optWM.Value = "W" Then
                    If dteEnding.Value & "" = "" Then
                        EMsg &= vbCr & "You Must Specify a Week-Ending Date"
                    Else
                        If CDate(dteEnding.Value).DayOfWeek <> DayOfWeek.Friday Then
                            EMsg &= vbCr & "You Must Specify a Week-Ending Date (a Friday)"
                        End If
                    End If
                Else
                    Call Validate_Code("OPS_YYYYPP")
                End If

                If Val(numLAST.Value) < 1 Or Val(numLAST.Value) > 13 Then
                    EMsg &= "Range must be between 1 and 13"
                End If

                selectedPriceCatgyCode = String.Empty
                selectedProdCatgyCode = String.Empty
                selectedVendor = String.Empty

                If MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Text.Trim.Length > 0 Then
                    Dim PRICE_CATGY_CODE As String = MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Text.Trim
                    Dim rowICTPCAT1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTPCAT1 WHERE  PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'")
                    If rowICTPCAT1 Is Nothing Then
                        EMsg &= "Invalid Price Category Code"
                    Else
                        selectedPriceCatgyCode = rowICTPCAT1.Item("PRICE_CATGY_CODE")
                        selectedProdCatgyCode = rowICTPCAT1.Item("PROD_CATGY_CODE")
                        selectedVendor = rowICTPCAT1.Item("VEND_CODE")
                    End If
                End If

            Case "Excel"
                If tabMain.ActiveTab.Key = "Summary" Then
                    If tabSummary.ActiveTab.Key = "Trend" Then
                        If grdICTINVAT.Rows.Count > 32000 Then
                            EMsg = "Excel has a 32,000 row limit."
                        End If
                    Else
                        If grdICTINVA0.Rows.Count > 32000 Then
                            EMsg = "Excel has a 32,000 row limit."
                        End If
                    End If

                ElseIf tabMain.ActiveTab.Key = "Items" Then
                    If grdICTINVA1.Rows.Count > 32000 Then
                        EMsg = "Excel has a 32,000 row limit."
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Call Set_Read_Only(UltraGroupBox1, True)
                Call Load_Record()

                If selectedProdCatgyCode.Length > 0 Then
                    viewICTINVAT.RowFilter = "CODE_VALUE = '" & selectedProdCatgyCode & "'"
                Else
                    viewICTINVAT.RowFilter = String.Empty
                End If
                grdICTINVAT.DataSource = viewICTINVAT

                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Excel"
                If tabMain.ActiveTab.Key = "Summary" Then
                    If tabSummary.ActiveTab.Key = "Trend" Then
                        Call Export_to_Excel(grdICTINVAT)
                    Else
                        Call Export_to_Excel(grdICTINVA0)
                    End If

                ElseIf tabMain.ActiveTab.Key = "Items" Then
                    Call Export_to_Excel(grdICTINVA1)
                End If
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Excel").Visible = False

                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode

                .Groups("Data").Visible = False
                .Groups("Show").Visible = False
                .Groups("Show Trend for").Visible = False
                .Groups("Other Options").Visible = False
                .Groups("Samples").Visible = False

            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = tf
        lblLoading.Visible = Not tf

        If ScreenMode Then
            tabMain.SelectedTab = tabMain.Tabs("Summary")
            Call Setup_tabMain()
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        chkShowSelectionPane.Checked = False
        chkChart.Checked = False

        'dst.EnforceConstraints = False
        MyBase.EnforceConstraints(False)
        dst.Tables("ICTINVA0").Rows.Clear()
        dst.Tables("ICTINVA1").Rows.Clear()
        dst.Tables("SOTINVHX").Rows.Clear()
        dst.Tables("POTORDRX").Rows.Clear()
        dst.Tables("SOTORDRX").Rows.Clear()
        dst.Tables("ICTINVAT").Rows.Clear()
        dst.Tables("APTVEND1").Rows.Clear()
        dst.Tables("ICTCATG1").Rows.Clear()
        dst.Tables("ICTPCAT1").Rows.Clear()
        'dst.EnforceConstraints = True
        MyBase.EnforceConstraints(True)

        MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Clear()

        viewICTINVAT.RowFilter = String.Empty

        selectedPriceCatgyCode = String.Empty
        selectedProdCatgyCode = String.Empty
        selectedVendor = String.Empty

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Inventory Analysis Data")
        Application.DoEvents()

        Dim PRICE_CATGY_CODE As String = MyBase.Absx1.txtFor("PRICE_CATGY_CODE").Text

        Call Save_Header_Fields(UltraGroupBox1)

        MOSMAX = Val(numLAST.Value)

        If optWM.Value = "W" Then
            grdICTINVA0.DisplayLayout.Bands(0).Columns("MOS").Header.Caption = "Wks"
            grdICTINVA1.DisplayLayout.Bands(0).Columns("MOS").Header.Caption = "Wks"
        Else
            grdICTINVA0.DisplayLayout.Bands(0).Columns("MOS").Header.Caption = "Mos"
            grdICTINVA1.DisplayLayout.Bands(0).Columns("MOS").Header.Caption = "Mos"
        End If

        'If ICTINVA1.Length = 0 Then
        ASCMAIN1.sql = "Select * from ICTINVA1 where ROWNUM < 1"
        ICTINVA1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add Primary Key (ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTINVA1 & "_1 on " & ICTINVA1 & " (PRICE_CATGY_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTINVA1 & "_2 on " & ICTINVA1 & " (VEND_CODE,PRICE_CATGY_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTINVA1 & "_3 on " & ICTINVA1 & " (PROD_CATGY_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTINVA1 & "_4 on " & ICTINVA1 & " (PROD_CATGY_CODE,VEND_CODE,PRICE_CATGY_CODE)")

        For I As Integer = 1 To MOSMAX
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add S" & Format(I, "00") & " Number (15,2)")
        Next
        For I As Integer = 1 To MOSMAX
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add L" & Format(I, "00") & " Number (8,0)")
        Next
        For I As Integer = 1 To MOSMAX
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add Q" & Format(I, "00") & " Number (8,0)")
        Next
        'Else
        'ASCDATA1.ExecuteSQL("Truncate Table " & ICTINVA1)
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_UOM")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN PRICE_CATGY_COST_TOTAL")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN PO_QTY_MULT")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_UPC_CODE")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_OPC_CODE")

        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_BASE_CURVE")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_DIAMETER")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_SPHERE_POWER")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_CYLINDER")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_AXIS")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_COLOR")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_ADD_POWER")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_ADD_DOM_NON")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_LEFT_RIGHT")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_CENTER_THICKNESS")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_BLANK_SIZE")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_INDEX_REF")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " DROP COLUMN ITEM_OZD")
        'End If

        If optWM.Value = "W" Then
            RDT = Absx1.dteFor("ENDING_DATE").Value
            FDT = RDT.AddDays(-7 * MOSMAX + 1)
        Else
            Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
            'z = Mid(z, 1, 4) & Mid(z, 6, 2)
            RYP = z
            FYP = ASCMAIN1.Period_Calc(RYP, 1 - MOSMAX)
        End If


        Dim ICTINVA1_X As String = ASCMAIN1.Temp_Table("Select * from " & ICTINVA1)
        ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1_X & " Add Primary Key (ITEM_CODE)")

        Dim ILIST As String = ""

        If "" = "X" Then
            ASCDATA1.ExecuteSQL("Insert into " & ICTINVA1 & " Select * from ICTINVA1")
        Else
            Dim sql As String = "Select ITEM_CODE" & vbCrLf
            For Each D As String In New String() {"S", "L", "Q"}
                Dim z As String = ""
                If D = "S" Then
                    z = "NVL(ORDR_QTY_SHIP,0) * NVL(ORDR_UNIT_PRICE,0)"
                ElseIf D = "L" Then
                    z = "1"
                ElseIf D = "Q" Then
                    z = "NVL(ORDR_QTY_SHIP,0)"
                End If
                For i As Integer = 1 To MOSMAX
                    If optWM.Value = "W" Then
                        Dim D2 As Date = RDT.AddDays(-7 * (MOSMAX - i))
                        Dim D1 As Date = D2.AddDays(-6)
                        sql &= ", SUM (CASE WHEN INV_DATE BETWEEN '" & Format(D1, "dd-MMM-yyyy") & "' AND '" & Format(D2, "dd-MMM-yyyy") & "' THEN " & z & " ELSE 0 END) "
                    Else
                        Dim YP As String = ASCMAIN1.Period_Calc(RYP, i - MOSMAX)
                        sql &= ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YP & "'," & z & ",0)) "
                    End If
                    sql &= D & Format(i, "00") & vbCrLf
                    ILIST &= "," & D & Format(i, "00")
                Next
            Next


            sql &= " from SOTINVH2" & vbCrLf
            If optWM.Value = "W" Then
                sql &= " where INV_DATE >= '" & Format(FDT, "dd-MMM-yyyy") & "' AND INV_DATE <= '" & Format(RDT, "dd-MMM-yyyy") & "'" & vbCrLf
            Else
                sql &= " where ORDR_YYYYPP_UPDATED >= '" & FYP & "' AND ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf
            End If

            If PRICE_CATGY_CODE.Length > 0 Then
                sql &= " and SOTINVH2.PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "' " & vbCrLf
            End If

            sql &= " group by ITEM_CODE"

            Call ASCMAIN1.Progress("Now Loading Sales History")

            sql = "Insert into " & ICTINVA1_X & "(ITEM_CODE" & ILIST & ") " & sql
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            sql = "INSERT INTO " & ICTINVA1_X & " (ITEM_CODE)"
            sql &= " SELECT ITEM_CODE FROM " & IIf(PRICE_CATGY_CODE.Length = 0, "ICTITEM1", "ICTCATL1") & " WHERE ITEM_ORDER_CODE IN ('N','P')"
            If PRICE_CATGY_CODE.Length > 0 Then
                sql &= " AND PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'"
            End If
            sql &= " MINUS"
            sql &= " SELECT ITEM_CODE FROM " & ICTINVA1_X

            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.AnalyzeTable(ICTINVA1_X)

            Call ASCMAIN1.Progress("Now Loading Item Status Data")

            sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS SELECT * FROM ICTSTAT2; " _
            & "BEGIN FOR R1 IN C1 LOOP " _
            & "UPDATE " & ICTINVA1_X & " SET " _
            & "  QOH = NVL(QOH,0) + NVL(R1.WHSE_QTY_ON_HAND,0)" _
            & ", QPO = NVL(QPO,0) + NVL(R1.WHSE_QTY_ONPO,0) " _
            & "WHERE ITEM_CODE = R1.ITEM_CODE; " _
            & "IF SQL%NOTFOUND THEN " _
            & "INSERT INTO " & ICTINVA1_X & " (ITEM_CODE, QOH, QPO) " _
            & "VALUES (R1.ITEM_CODE, NVL(R1.WHSE_QTY_ON_HAND,0), NVL(R1.WHSE_QTY_ONPO,0)); " _
            & "END IF; " _
            & "END LOOP; END; END; "
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            'Call ASCMAIN1.Progress("Now Calculating 1st/Last Sale Data")

            'sql = "SELECT ITEM_CODE " _
            '& ", MIN (INV_DATE) D1, MAX (INV_DATE) D2 " _
            '& ", MIN (ORDR_YYYYPP_UPDATED) P1, MAX (ORDR_YYYYPP_UPDATED) P2 " _
            '& " FROM SOTINVH2 GROUP BY ITEM_CODE "
            'Dim ICTINVA1_M As String = ASCMAIN1.Temp_Table(sql)
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1_M & " Add Primary Key (ITEM_CODE)")

            Call ASCMAIN1.Progress("Now Loading Item Master Data")

            Call ASCMAIN1.AnalyzeTable(ICTINVA1_X)
            Application.DoEvents()

            Dim sqlICTITEM1 As String = "ICTITEM1"
            If PRICE_CATGY_CODE.Length > 0 Then
                sqlICTITEM1 = "SELECT ITEM_CODE, ITEM_STATUS, ITEM_DESC, PRICE_CATGY_CODE, ITEM_BIN, ITEM_ORDER_CODE, ITEM_MIN_QTY, ITEM_MAX_QTY"
                sqlICTITEM1 &= " FROM ICTITEM1 WHERE PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'"
                sqlICTITEM1 &= " UNION"
                sqlICTITEM1 &= " SELECT ITEM_CODE, 'A' ITEM_STATUS, ITEM_DESC, PRICE_CATGY_CODE, NULL ITEM_BIN, ITEM_ORDER_CODE, NULL ITEM_MIN_QTY, NULL ITEM_MAX_QTY"
                sqlICTITEM1 &= " FROM ICTCATL1 WHERE PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'"
                sqlICTITEM1 &= " AND ITEM_CODE NOT IN (SELECT ITEM_CODE FROM ICTITEM1 WHERE PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "')"
                sqlICTITEM1 = "( " & sqlICTITEM1 & " ) ICTITEM1"
            End If

            sql = "Insert into " & ICTINVA1 & " Select X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_STATUS, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.PRICE_CATGY_CODE, ICTPCAT1.VEND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_BIN, ICTITEM1.ITEM_ORDER_CODE" & vbCrLf _
            & ", NULL PROD_CATGY_CODE, NULL PRICE_CATGY_SAMPLE_IND" & vbCrLf _
            & ", ICTITEM1.ITEM_MIN_QTY, ICTITEM1.ITEM_MAX_QTY" & vbCrLf _
            & ", ICTITEM5.INIT_SOLD, ICTITEM5.INIT_SOLD_YYYYPP" & vbCrLf _
            & ", ICTITEM5.LAST_SOLD, ICTITEM5.LAST_SOLD_YYYYPP" & vbCrLf _
            & ", X.QOH, X.QPO, X.QSO, NULL MOS, NULL S00, NULL L00, NULL Q00" & vbCrLf _
            & ILIST & vbCrLf _
            & " from " & ICTINVA1_X & " X, ICTITEM5, " & sqlICTITEM1 & " , ICTPCAT1" & vbCrLf _
            & " where X.ITEM_CODE = ICTITEM1.ITEM_CODE" & vbCrLf _
            & " and ICTPCAT1.PRICE_CATGY_CODE (+) = ICTITEM1.PRICE_CATGY_CODE" & vbCrLf _
            & " and ICTITEM5.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            sql_ICTITEMX = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & ILIST & " from " & ICTINVA1 & " ICTINVA1, ICTITEM1 where ICTINVA1.ITEM_CODE = ICTITEM1.ITEM_CODE and ICTINVA1.ITEM_CODE = :PARM1"

            Call ASCMAIN1.Progress("Now Loading Price Category Data")

            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_UOM VARCHAR2(4)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add PRICE_CATGY_COST_TOTAL NUMBER (8,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add PO_QTY_MULT Number (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_UPC_CODE VARCHAR2(18)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_OPC_CODE VARCHAR2(10)")
            Application.DoEvents()

            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_BASE_CURVE NUMBER (8, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_DIAMETER NUMBER( 9, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_SPHERE_POWER NUMBER (9, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_CYLINDER NUMBER (6, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_AXIS NUMBER (4, 0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_COLOR VARCHAR2 (3)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_ADD_POWER VARCHAR2 (4)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_ADD_DOM_NON VARCHAR2 (1)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_LEFT_RIGHT VARCHAR2 (1)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_CENTER_THICKNESS NUMBER (6, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_BLANK_SIZE NUMBER (6, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_INDEX_REF NUMBER (8, 2)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTINVA1 & " Add ITEM_OZD NUMBER (6, 2)")
            Application.DoEvents()

            sql = "" _
            & " BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT * FROM ICTPCAT1;" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & " UPDATE " & ICTINVA1 & " SET VEND_CODE = R1.VEND_CODE" _
            & ", PROD_CATGY_CODE = R1.PROD_CATGY_CODE " _
            & ", PRICE_CATGY_SAMPLE_IND = NVL(R1.PRICE_CATGY_SAMPLE_IND,'0') " _
            & ", PO_QTY_MULT = NVL(R1.PO_QTY_MULT, 1) " _
            & " WHERE PRICE_CATGY_CODE = R1.PRICE_CATGY_CODE; " _
            & " END LOOP; END; END; "
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Loading Cost Data")
            sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT * FROM ICTCOSTC;" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & " UPDATE " & ICTINVA1 & " SET PRICE_CATGY_COST_TOTAL = R1.PRICE_CATGY_COST_TOTAL" _
            & " WHERE PRICE_CATGY_CODE = R1.PRICE_CATGY_CODE; " _
            & " END LOOP; END; END; "
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Loading Uom")
            'sql = "UPDATE " & ICTINVA1 & " SET ITEM_UOM = (SELECT ITEM_UOM FROM ICTITEM1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE)"
            sql = "UPDATE " & ICTINVA1 & " SET ITEM_UOM = (SELECT ITEM_UOM FROM ICTCATL1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE)"
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Loading UPCs")
            sql = "UPDATE " & ICTINVA1 & " SET ITEM_UPC_CODE = (SELECT ITEM_UPC_CODE FROM ICTITEM1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE)"
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Loading UPCs")
            sql = "UPDATE " & ICTINVA1 & " SET ITEM_UPC_CODE = (SELECT ITEM_PROD_ID FROM ICTCATL1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE AND ITEM_PROD_QUAL = 'UP') WHERE ITEM_UPC_CODE IS NULL"
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()


            Call ASCMAIN1.Progress("Now Loading OPCs")
            sql = "UPDATE " & ICTINVA1 & " SET ITEM_OPC_CODE = (SELECT ITEM_OPC_CODE FROM ICTITEM1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE)"
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Loading OPCs")
            sql = "UPDATE " & ICTINVA1 & " SET ITEM_OPC_CODE = (SELECT ITEM_PROD_ID FROM ICTCATL1 WHERE ITEM_CODE = " & ICTINVA1 & ".ITEM_CODE AND ITEM_PROD_QUAL = 'OI') WHERE ITEM_OPC_CODE IS NULL"
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()
            Application.DoEvents()


            Call ASCMAIN1.Progress("Now Loading Parameter Data")
            sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT ITEM_CODE FROM " & ICTINVA1 & ");" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & " UPDATE " & ICTINVA1 & " SET " _
            & "   ITEM_BASE_CURVE = R1.ITEM_BASE_CURVE" _
            & " , ITEM_SPHERE_POWER = R1.ITEM_SPHERE_POWER" _
            & " , ITEM_CYLINDER = R1.ITEM_CYLINDER" _
            & " , ITEM_AXIS = R1.ITEM_AXIS" _
            & " , ITEM_COLOR = R1.ITEM_COLOR" _
            & " , ITEM_ADD_POWER = R1.ITEM_ADD_POWER" _
            & " , ITEM_ADD_DOM_NON = R1.ITEM_ADD_DOM_NON" _
            & " , ITEM_LEFT_RIGHT = R1.ITEM_LEFT_RIGHT" _
            & " , ITEM_CENTER_THICKNESS = R1.ITEM_CENTER_THICKNESS" _
            & " , ITEM_BLANK_SIZE = R1.ITEM_BLANK_SIZE" _
            & " , ITEM_INDEX_REF = R1.ITEM_INDEX_REF" _
            & " , ITEM_OZD = R1.ITEM_OZD" _
            & " WHERE ITEM_CODE = R1.ITEM_CODE; " _
            & " END LOOP; END; END; "
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

            Call ASCMAIN1.Progress("Now Calculating Totals")

            Dim sqlx As String = ""
            For i As Integer = 1 To MOSMAX
                sqlx &= "+NVL(X" & Format(i, "00") & ",0)"
            Next
            sqlx = Mid(sqlx, 2)

            'sql = "UPDATE " & ICTINVA1 & " SET S00 = " & Replace(sqlx, "(X", "(S")
            'ASCDATA1.ExecuteSQL(sql)
            'sql = "UPDATE " & ICTINVA1 & " SET L00 = " & Replace(sqlx, "(X", "(L")
            'ASCDATA1.ExecuteSQL(sql)
            'sql = "UPDATE " & ICTINVA1 & " SET Q00 = " & Replace(sqlx, "(X", "(Q")
            'ASCDATA1.ExecuteSQL(sql)
            'sql = "UPDATE " & ICTINVA1 & " SET MOS  = " & Replace(Replace(sqlx, "NVL(X", "SIGN(ABS(L"), ",0)", "))")
            'ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE " & ICTINVA1 & " Set " _
            & "  S00 = " & Replace(sqlx, "(X", "(S") _
            & ", L00 = " & Replace(sqlx, "(X", "(L") _
            & ", Q00 = " & Replace(sqlx, "(X", "(Q") _
            & ", MOS = " & Replace(Replace(sqlx, "NVL(X", "SIGN(ABS(L"), ",0)", "))")
            ASCDATA1.ExecuteSQL(sql)
            Application.DoEvents()

        End If

        Call ASCMAIN1.Progress("Now Analyzing Tables")

        Call ASCMAIN1.AnalyzeTable(ICTINVA1)

        ASCDATA1.ExecuteSQL("Update " & ICTINVA1 & " Set MOS = Null where MOS = 0")
        Application.DoEvents()

        Call ASCMAIN1.Progress("Now Loading Data")

        'dst.EnforceConstraints = False
        MyBase.EnforceConstraints(False)

        Call Fill_Records("ICTCATG1", "", , "Select ICTCATG1.PROD_CATGY_CODE,ICTCATG1.PROD_CATGY_DESC, Count (*) ITEMS from " & ICTINVA1 & " ICTINVA1, ICTCATG1 where ICTINVA1.PROD_CATGY_CODE = ICTCATG1.PROD_CATGY_CODE group by ICTCATG1.PROD_CATGY_CODE,ICTCATG1.PROD_CATGY_DESC")
        Call Load_VEND_CODEs()
        Call Load_PRICE_CATGY_CODEs()

        'dst.EnforceConstraints = True
        MyBase.EnforceConstraints(True)

        Call ASCMAIN1.Progress("Now Setting Up Screen")

        With dst.Tables("ICTINVA1").Columns
            For i As Integer = .Count - 1 To 0 Step -1
                Dim X_COLUMN_NAME As String = .Item(i).ColumnName
                If Len(X_COLUMN_NAME) = 3 And Mid(X_COLUMN_NAME, 2, 2) <> "00" _
                And X_COLUMN_NAME <> "QOH" _
                And X_COLUMN_NAME <> "QPO" _
                And X_COLUMN_NAME <> "QSO" _
                And (X_COLUMN_NAME Like "S*" Or _
                     X_COLUMN_NAME Like "L*" Or _
                     X_COLUMN_NAME Like "Q*") Then
                    .Remove(X_COLUMN_NAME)
                    With grdICTINVA1.DisplayLayout.Bands(0)
                        .Summaries.Remove(.Summaries(X_COLUMN_NAME))
                        ' .Columns.Remove(X_COLUMN_NAME)
                    End With
                End If
            Next

            With grdICTINVA1.DisplayLayout.Bands(0)
                .Groups.Remove("S")
                .Groups.Remove("L")
                .Groups.Remove("Q")
                .Groups.Add("S", "Sales")
                .Groups.Add("L", "Line Items")
                .Groups.Add("Q", "Qty Shipped")
            End With

            For Each slq As String In New String() {"S", "L", "Q"}


                For i As Integer = 1 To MOSMAX
                    Dim X_COLUMN_NAME As String = slq & Format(i, "00")
                    If slq = "S" Then
                        .Add(X_COLUMN_NAME, GetType(System.Double))
                    Else
                        .Add(X_COLUMN_NAME, GetType(System.Int32))
                    End If
                    With grdICTINVA1.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME)
                        .Width = grdICTINVA1.DisplayLayout.Bands(0).Columns(slq & "00").Width
                        .Format = grdICTINVA1.DisplayLayout.Bands(0).Columns(slq & "00").Format
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right

                        .Group = grdICTINVA1.DisplayLayout.Bands(0).Groups(slq)
                        .Hidden = False
                    End With

                    Call Create_Summary(grdICTINVA1, slq & Format(i, "00"))
                Next
            Next
        End With

        With grdICTINVA1.DisplayLayout.Bands(0)
            .Groups("Item Attributes").CellAppearance.BackColor = Color.Beige
            .Groups("Totals").CellAppearance.BackColor = Color.LightSkyBlue
            .Groups("S").CellAppearance.BackColor = Color.LightGreen
            .Groups("L").CellAppearance.BackColor = Color.LightSalmon
            .Groups("Q").CellAppearance.BackColor = Color.AliceBlue
        End With

        With grdICTINVA1.DisplayLayout.Bands(0)
            For i As Integer = 1 To MOSMAX
                Dim z As String = ""
                If optWM.Value = "W" Then
                    z = Format(FDT.AddDays(6 + (i - 1) * 7), "MM/dd")
                Else
                    z = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(FYP, i - 1))
                    z = Mid(z, 10, 6)
                End If
                .Columns("S" & Format(i, "00")).Header.Caption = z
                .Columns("L" & Format(i, "00")).Header.Caption = z
                .Columns("Q" & Format(i, "00")).Header.Caption = z
            Next
        End With

        Call Setup_Columns()

        Call Fetch()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdICTINVA0, "OOO", _
                            "OnlyOneUp:All Items in and above Selected Cell", _
                            "OnlyOneUp:All Items in and below Selected Cell", _
                            "OnlyOneUp:Show All Items in Selected Cell Only")
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        If Not Me.Visible Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub

        Select Case e.Tool.Key
            Case "All Items in and above Selected Cell" ' Selected Cell only"
                optCell.CheckedIndex = 2
            Case "All Items in and below Selected Cell" ' All Cells <= Selected Cell"
                optCell.CheckedIndex = 1
            Case "Show All Items in Selected Cell Only" ' "All Cells >= Selected Cell"
                optCell.CheckedIndex = 0
        End Select
    End Sub

#End Region


    Private Sub chkSHOWI_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWI.CheckedChanged
        Call Setup_Columns()
    End Sub

    Private Sub chkSHOWS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWS.CheckedChanged
        Call Setup_Columns()
    End Sub

    Private Sub chkSHOWL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWL.CheckedChanged
        Call Setup_Columns()
    End Sub

    Private Sub chkSHOWP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWP.CheckedChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        grdICTINVA1.DisplayLayout.Bands(0).Groups("Item Parameters").Hidden = Not chkSHOWP.Checked
    End Sub

    Private Sub chkSHOWQ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWQ.CheckedChanged
        Call Setup_Columns()
    End Sub

    Sub Setup_Columns()
        If dst.Tables.Count = 0 Then Exit Sub

        With grdICTINVA0.DisplayLayout.Bands(0)
            .Groups("I").Hidden = Not chkSHOWI.Checked
            .Groups("S").Hidden = Not chkSHOWS.Checked
            .Groups("L").Hidden = Not chkSHOWL.Checked
            .Groups("Q").Hidden = Not chkSHOWQ.Checked

            .Groups("S").Width = .Groups("I").Width
            .Groups("L").Width = .Groups("I").Width
            .Groups("Q").Width = .Groups("I").Width

            For Each T As String In New String() {"ALL", "STK", "RXP", "FRZ"}
                .Columns("S_" & T).Width = .Columns("I_" & T).Width
                .Columns("L_" & T).Width = .Columns("I_" & T).Width
                .Columns("Q_" & T).Width = .Columns("I_" & T).Width
            Next
        End With

        With grdICTINVA1.DisplayLayout.Bands(0)
            .Groups("S").Hidden = Not chkSHOWS.Checked
            .Groups("L").Hidden = Not chkSHOWL.Checked
            .Groups("Q").Hidden = Not chkSHOWQ.Checked
        End With

    End Sub

    Private Sub cmdFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFetch.Click
        Call Fetch()
    End Sub

    Sub Fetch_if_Auto()
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub

        If chkAutoFetch.Checked Then
            Call Fetch()
        End If
    End Sub

    Sub Fetch()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Fetching Data")
        Application.DoEvents()

        Dim filters As String = ""

        Dim sql As String = "Select ICTINVA1.MOS"


        For Each D As String In New String() {"I", "S", "L", "Q"}
            Dim z As String
            If D = "I" Then
                z = "1"
            Else
                z = "ICTINVA1." & D & "00"
            End If
            sql &= ", SUM (" & z & ") " & D & "_ALL"
            sql &= ", SUM (DECODE(ICTINVA1.ITEM_ORDER_CODE,'N'," & z & ",0)) " & D & "_STK"
            sql &= ", SUM (DECODE(ICTINVA1.ITEM_ORDER_CODE,'P'," & z & ",0)) " & D & "_RXP"
            sql &= ", SUM (DECODE(ICTINVA1.ITEM_ORDER_CODE,'X'," & z & ",0)) " & D & "_FRZ"
        Next

        sql &= " from " & ICTINVA1 & " ICTINVA1"

        sqlw = ""

        If optASR.Value = "S" Then
            sqlw &= " and NVL(ICTINVA1.PRICE_CATGY_SAMPLE_IND,'0') = '1'"
            filters &= ", (Samples Only)"
        ElseIf optASR.Value = "R" Then
            sqlw &= " and NVL(ICTINVA1.PRICE_CATGY_SAMPLE_IND,'0') <> '1'"
            filters &= ", (Revenue Only)"
        End If

        If chkSalesOnly.Checked Then
            sqlw &= " and NVL(ICTINVA1.S00,0) <> 0"
            filters &= ", with Sales"
        End If
        If chkOHOnly.Checked Then
            sqlw &= " and NVL(ICTINVA1.QOH,0) <> 0"
            filters &= ", with On-Hand"
        End If

        If chkExcludeNew.Checked Then
            filters &= ", excluding New Launches"
            If optWM.Value = "W" Then
                sqlw &= " and ICTINVA1.FST_SALE_DATE <= '" & Format(FDT, "dd-MMM-yyyy") & "'"
            Else
                sqlw &= " and ICTINVA1.FST_SALE_YYYYPP <= '" & FYP & "'"
            End If
        End If

        If grdICTCATG1.Selected.Rows.Count <> 0 And optPROD_CATGY_CODE.Value = "S" Then
            Dim z As String = ""
            Dim ZV As String = ""
            For Each grow As UltraWinGrid.UltraGridRow In grdICTCATG1.Selected.Rows
                z = z & ",'" & grow.Cells("PROD_CATGY_CODE").Text & "'"
                ZV = ZV & "," & grow.Cells("PROD_CATGY_CODE").Text
            Next
            If grdICTCATG1.Selected.Rows.Count = 1 Then
                filters &= ", Product Catgy " & Mid(ZV, 2)
                sqlw &= " and ICTINVA1.PROD_CATGY_CODE = " & Mid(z, 2)
            Else
                filters &= ", Selected Product Catgys"
                sqlw &= " and ICTINVA1.PROD_CATGY_CODE in (" & Mid(z, 2) & ")"
            End If
        End If

        If grdAPTVEND1.Selected.Rows.Count <> 0 And optVEND_CODE.Value = "S" Then
            Dim z As String = ""
            Dim ZV As String = ""
            For Each grow As UltraWinGrid.UltraGridRow In grdAPTVEND1.Selected.Rows
                z = z & ",'" & grow.Cells("VEND_CODE").Text & "'"
                ZV = ZV & "," & grow.Cells("VEND_CODE").Text
            Next
            If grdAPTVEND1.Selected.Rows.Count = 1 Then
                filters &= ", Vendor " & Mid(ZV, 2)
                sqlw &= " and ICTINVA1.VEND_CODE = " & Mid(z, 2)
            Else
                filters &= ", Selected Vendors"
                sqlw &= " and ICTINVA1.VEND_CODE in (" & Mid(z, 2) & ")"
            End If
        End If

        If grdICTPCAT1.Selected.Rows.Count <> 0 And optPRICE_CATGY_CODE.Value = "S" Then
            Dim z As String = ""
            Dim ZV As String = ""
            For Each grow As UltraWinGrid.UltraGridRow In grdICTPCAT1.Selected.Rows
                z = z & ",'" & grow.Cells("PRICE_CATGY_CODE").Text & "'"
                ZV = ZV & "," & grow.Cells("PRICE_CATGY_CODE").Text
            Next
            If grdICTPCAT1.Selected.Rows.Count = 1 Then
                filters &= ", Price Catgy " & Mid(ZV, 2)
                sqlw &= " and ICTINVA1.PRICE_CATGY_CODE = " & Mid(z, 2)
            Else
                filters &= ", Selected Price Catgys"
                sqlw &= " and ICTINVA1.PRICE_CATGY_CODE in (" & Mid(z, 2) & ")"
            End If
        End If

        If filters = "" Then
            grdICTINVA0.Text = "Including All Active Items"
        Else
            grdICTINVA0.Text = "Including Active Items" & Mid$(filters, 2)
        End If

        grdICTINVAT.Text = grdICTINVA0.Text

        sql &= ASCMAIN1.SQL_Add_WHERE(sqlw) & " group by ICTINVA1.MOS"

        Call Fill_Records("ICTINVA0", "", , sql)
        tabMain.Tabs("Items").Enabled = False

        If grdICTINVA0.Rows.Count > 0 Then
            For i As Integer = 0 To grdICTINVA0.Rows.Count - 1
                If grdICTINVA0.Rows(i).Cells(0).Text = "" Then
                    grdICTINVA0.Rows(i).ToolTipText = "This row shows Items which have a non-zero Qty On Hand or PO Qty, but haven't shipped in the last " & CStr(MOSMAX) & " " & grdICTINVA0.DisplayLayout.Bands(0).Columns("MOS").Header.Caption
                Else
                    grdICTINVA0.Rows(i).ToolTipText = "This row shows Items which have shipped in " & grdICTINVA0.Rows(i).Cells(0).Text & " of the last " & CStr(MOSMAX) & " " & grdICTINVA0.DisplayLayout.Bands(0).Columns("MOS").Header.Caption
                End If
            Next
        End If

        If grdICTINVA0.Rows.Count <> 0 Then
            grdICTINVA0.ActiveRow = grdICTINVA0.Rows(0)
        End If
        If grdICTINVA1.Rows.Count <> 0 Then
            grdICTINVA1.ActiveRow = grdICTINVA1.Rows(0)
        End If

        Call Fetch_Trend()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()

    End Sub

    Private Sub grdICTINVA0_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdICTINVA0.DoubleClickCell

        If grdICTINVA0.ActiveCell Is Nothing Then
            Exit Sub
        End If

        Dim COL As String = grdICTINVA0.ActiveCell.Column.Key
        If Len(COL) = 5 And Mid(COL, 2, 1) = "_" Then
            Call Setup_ICTINVA1(COL, Val(grdICTINVA0.ActiveRow.Cells("MOS").Text))
        End If

    End Sub

    Private Sub grdICTINVA0_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTINVA0.DoubleClickRow
        'Stop
    End Sub

    Private Sub grdICTINVA0_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTINVA0.InitializeLayout
        With grdICTINVA0.DisplayLayout.Bands(0)
            .Groups.Add("M", "")
            .Groups.Add("I", "Item Count")
            .Groups.Add("S", "Sales")
            .Groups.Add("L", "Line Items")
            .Groups.Add("Q", "Qty Shipped")

            .Groups("I").CellAppearance.BackColor = Color.Beige
            .Groups("S").CellAppearance.BackColor = Color.LightGreen
            .Groups("L").CellAppearance.BackColor = Color.LightSalmon
            .Groups("Q").CellAppearance.BackColor = Color.AliceBlue

            For Each c As UltraWinGrid.UltraGridColumn In .Columns
                If Len(c.Key) = 5 And Mid(c.Key, 2, 1) = "_" Then
                    c.Header.Caption = ASCMAIN1.Make_Caption(Mid(c.Key, 3, 3))
                    c.Group = .Groups(Mid(c.Key, 1, 1))
                Else
                    c.Header.Caption = ASCMAIN1.Make_Caption(c.Key)
                    c.Group = .Groups("M")
                End If
            Next
        End With
    End Sub

    Private Sub optWM_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optWM.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub

        dteEnding.Visible = (optWM.Value = "W")
        Absx1.txtFor("OPS_YYYYPP").Visible = (optWM.Value = "M")
        Absx1.txtFor("LEGEND").Visible = (optWM.Value = "M")
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If dst.Tables.Count = 0 Then Exit Sub
        Call Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        UltraExplorerBar1.Groups("Data").Visible = False
        UltraExplorerBar1.Groups("Show").Visible = False
        UltraExplorerBar1.Groups("Show Trend for").Visible = False
        UltraExplorerBar1.Groups("Other Options").Visible = False
        UltraExplorerBar1.Groups("Samples").Visible = False

        Select Case tabMain.SelectedTab.Key
            Case "Summary"
                Call Setup_tabSummary()

            Case "Items"
                UltraExplorerBar1.Groups("Show").Visible = True
                chkSHOWI.Enabled = False
                chkSHOWA.Enabled = True
        End Select
    End Sub

    Sub Setup_tabSummary()
        'UltraExplorerBar1.Groups("Data").Visible = True
        UltraExplorerBar1.Groups("Other Options").Visible = True
        UltraExplorerBar1.Groups("Samples").Visible = True
        chkSHOWI.Enabled = True
        chkSHOWA.Enabled = False
        UltraExplorerBar1.Groups("Show Trend for").Visible = False
        UltraExplorerBar1.Groups("Show").Visible = False

        Select Case tabSummary.SelectedTab.Key
            Case "Trend"
                UltraExplorerBar1.Groups("Show Trend for").Visible = True
            Case "Analysis"
                UltraExplorerBar1.Groups("Show").Visible = True
        End Select
    End Sub

    Private Sub chkShowSelectionPane_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowSelectionPane.CheckedChanged
        SplitContainer2.Panel2Collapsed = Not chkShowSelectionPane.Checked
    End Sub

    Private Sub cmdPROD_CATGY_CODE_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPROD_CATGY_CODE_Clear.Click
        optPROD_CATGY_CODE.Value = "A"
        grdICTCATG1.Selected.Rows.Clear()
    End Sub

    Private Sub cmdVEND_CODE_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVEND_CODE_Clear.Click
        optVEND_CODE.Value = "A"
        grdAPTVEND1.Selected.Rows.Clear()
    End Sub

    Private Sub grdICTPCAT1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPCAT1.AfterRowActivate

    End Sub

    Private Sub grdICTPCAT1_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdICTPCAT1.AfterSelectChange
        If optPRICE_CATGY_CODE.Value = "S" Then Call Fetch_if_Auto()
    End Sub

    Private Sub optPRICE_CATGY_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPRICE_CATGY_CODE.ValueChanged
        Call Fetch_if_Auto()
    End Sub

    Private Sub optVEND_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVEND_CODE.ValueChanged
        Call Fetch_if_Auto()
        chkW_VEND_CODE.Visible = (optVEND_CODE.Value = "S")
    End Sub

    Private Sub grdAPTVEND1_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdAPTVEND1.AfterSelectChange
        If optVEND_CODE.Value = "S" Then
            Call Fetch_if_Auto()
            If chkW_VEND_CODE.Checked Then
                Call Load_PRICE_CATGY_CODEs()
            End If
        End If
    End Sub

    Private Sub optPROD_CATGY_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPROD_CATGY_CODE.ValueChanged
        Call Fetch_if_Auto()
        chkW_PROD_CATGY_CODE.Visible = (optPROD_CATGY_CODE.Value = "S")
    End Sub

    Private Sub grdICTCATG1_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdICTCATG1.AfterSelectChange
        If optPROD_CATGY_CODE.Value = "S" Then
            Call Fetch_if_Auto()
            If chkW_PROD_CATGY_CODE.Checked Then
                Call Load_VEND_CODEs()
            End If
        End If
    End Sub

    Private Sub chkSalesOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSalesOnly.CheckedChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_if_Auto()

    End Sub

    Private Sub chkOHOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOHOnly.CheckedChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_if_Auto()
    End Sub

    Private Sub chkExcludeNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExcludeNew.CheckedChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_if_Auto()
    End Sub

    Private Sub optASR_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optASR.ValueChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_if_Auto()
    End Sub

    Private Sub chkSHOWA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSHOWA.CheckedChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        grdICTINVA1.DisplayLayout.Bands(0).Groups("Item Attributes").Hidden = Not chkSHOWA.Checked
    End Sub

    Private Sub grdICTINVA1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdICTINVA1.DoubleClickCell

        If grdICTINVA1.ActiveCell Is Nothing Then
            Exit Sub
        End If

        Call Set_ITEM_CODE(grdICTINVA1.ActiveRow.Cells("ITEM_CODE").Text)
    End Sub

    Sub Set_ITEM_CODE(ByVal ITEM_CODE As String)

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Call ASCMAIN1.Progress("Now Loading Open PO Details - Item " & ITEM_CODE)
        Call Fill_Records("POTORDRX", ITEM_CODE)

        Call ASCMAIN1.Progress("Now Loading Open SO Details - Item " & ITEM_CODE)
        Call Fill_Records("SOTORDRX", ITEM_CODE)

        Call ASCMAIN1.Progress("Now Loading Sales Details - Item " & ITEM_CODE)

        Dim sql As String = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO, SOTINVH2.INV_DATE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_UNIT_PRICE_PATIENT from SOTINVH2,SOTINVH1,ARTCUST1 where ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO and SOTINVH2.ITEM_CODE = '" & ITEM_CODE & "'"
        'sql &= " and SOTINVH2.INV_TYPE = 'I'"

        Dim time_range_handled As Boolean = False

        Dim COL As String = grdICTINVA1.ActiveCell.Column.Key
        If Len(COL) = 3 Then
            Select Case COL
                Case "QPO"
                    tabDetails.SelectedTab = tabDetails.Tabs("Open Purchase Orders")
                Case "QSO"
                    tabDetails.SelectedTab = tabDetails.Tabs("Open Sales Orders")
                Case Else
                    tabDetails.SelectedTab = tabDetails.Tabs("Sales History Detail")

                    If InStr("SLQ", Mid(COL, 1, 1)) <> 0 And Mid(COL, 2, 2) <> "00" Then
                        Dim P As Integer = Val(Mid(COL, 2, 2))
                        If optWM.Value = "W" Then
                            sql &= " and SOTINVH2.INV_DATE >= '" & Format(FDT.AddDays((P - 1) * 7), "dd-MMM-yyyy") & "' and SOTINVH2.INV_DATE <= '" & Format(FDT.AddDays((P - 1) * 7 + 6), "dd-MMM-yyyy") & "'"
                        Else
                            sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.Period_Calc(FYP, P - 1) & "'"
                        End If

                        time_range_handled = True
                    End If
            End Select
        End If

        If Not time_range_handled Then
            If optWM.Value = "W" Then
                sql &= " and SOTINVH2.INV_DATE >= '" & Format(FDT, "dd-MMM-yyyy") & "' and SOTINVH2.INV_DATE <= '" & Format(RDT, "dd-MMM-yyyy") & "'"
            Else
                sql &= " and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & FYP & "' and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'"
            End If
        End If

        Call Fill_Records("SOTINVHX", "", , sql)

        tabDetails.Tabs("Item").Text = "Item " & ITEM_CODE
        Call Fill_Records("ICTITEMX", ITEM_CODE, , sql_ICTITEMX)

        SplitContainer6.Panel2Collapsed = False

        grdICTINVAI.DataSource = New DataView(dst.Tables("ICTINVA1"), "ITEM_CODE = '" & ITEM_CODE & "'", "", DataViewRowState.CurrentRows).ToTable
        With grdICTINVAI.DisplayLayout
            .Load(grdICTINVA1.DisplayLayout, UltraWinGrid.PropertyCategories.All)
            .GroupByBox.Hidden = True
            .Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.None
            .Bands(0).Groups("Item Attributes").Hidden = False
            .Bands(0).Groups("S").Hidden = False
            .Bands(0).Groups("L").Hidden = False
            .Bands(0).Groups("Q").Hidden = False
        End With
        grdICTINVAI.Text = ""


        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()

    End Sub

    Private Sub grdICTINVA1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTINVA1.DoubleClickRow

    End Sub

    Private Sub grdICTINVA1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTINVA1.InitializeLayout

    End Sub

    Private Sub chkW_PROD_CATGY_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkW_PROD_CATGY_CODE.CheckedChanged
        Call Load_VEND_CODEs()
    End Sub

    Private Sub chkW_VEND_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkW_VEND_CODE.CheckedChanged
        Call Load_PRICE_CATGY_CODEs()
    End Sub

    Sub Load_VEND_CODEs()
        grdAPTVEND1.Selected.Rows.Clear()
        optVEND_CODE.Value = "A"

        Dim sql As String = "Select APTVEND1.VEND_CODE,APTVEND1.VEND_NAME " & vbCr _
        & ", Count (*) ITEMS from " & ICTINVA1 & " ICTINVA1, APTVEND1 " & vbCr _
        & " where APTVEND1.VEND_CODE = ICTINVA1.VEND_CODE"
        If chkW_PROD_CATGY_CODE.Checked And optPROD_CATGY_CODE.Value = "S" And grdICTCATG1.Selected.Rows.Count <> 0 Then

            Dim z As String = ""
            For Each grow As UltraWinGrid.UltraGridRow In grdICTCATG1.Selected.Rows
                z = z & ",'" & grow.Cells("PROD_CATGY_CODE").Text & "'"
            Next

            sql &= " and ICTINVA1.PROD_CATGY_CODE in (" & Mid(z, 2) & ")"
        End If

        sql &= " group by APTVEND1.VEND_CODE, APTVEND1.VEND_NAME"
        Call Fill_Records("APTVEND1", "", , sql)
    End Sub

    Sub Load_PRICE_CATGY_CODEs()
        grdICTPCAT1.Selected.Rows.Clear()
        optPRICE_CATGY_CODE.Value = "A"

        Dim sql As String = "Select ICTPCAT1.PRICE_CATGY_CODE" & vbCr _
        & ", ICTPCAT1.PRICE_CATGY_DESC, ICTPCAT1.VEND_CODE" & vbCr _
        & ", Count (*) ITEMS " & vbCr _
        & " from " & ICTINVA1 & " ICTINVA1, ICTPCAT1 " & vbCr _
        & " where ICTINVA1.PRICE_CATGY_CODE = ICTPCAT1.PRICE_CATGY_CODE "

        If chkW_VEND_CODE.Checked And optVEND_CODE.Value = "S" And grdAPTVEND1.Selected.Rows.Count <> 0 Then

            Dim z As String = ""
            For Each grow As UltraWinGrid.UltraGridRow In grdAPTVEND1.Selected.Rows
                z = z & ",'" & grow.Cells("VEND_CODE").Text & "'"
            Next

            sql &= " and ICTINVA1.VEND_CODE in (" & Mid(z, 2) & ")"
        End If

        sql &= " group by ICTPCAT1.PRICE_CATGY_CODE" & vbCr _
        & ", ICTPCAT1.PRICE_CATGY_DESC, ICTPCAT1.VEND_CODE"
        Call Fill_Records("ICTPCAT1", "", , sql)
    End Sub

    Private Sub grdICTINVA0_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdICTINVA0.MouseDown

        Dim pt As Point = New Point(e.X, e.Y)
        Dim elem As Infragistics.Win.UIElement
        elem = grdICTINVA0.DisplayLayout.UIElement.ElementFromPoint(pt)
        If Not Nothing Is elem And Not Nothing Is elem.Parent Then
            Dim summaryElemType As System.Type = _
            GetType(Infragistics.Win.UltraWinGrid.SummaryValueUIElement)
            If elem.Parent.GetType().Equals(summaryElemType) Then
                'Stop

                'Dim sv As Infragistics.Win.UltraWinGrid.SummaryValue = _
                'DirectCast(elem.Parent, Infragistics.Win.UltraWinGrid.SummaryValue)
                '  Stop

                'Call Setup_ICTINVA1(col, mos)
            End If
        End If

    End Sub

    Sub Setup_ICTINVA1(ByVal COL As String, ByVal MOS As Integer, Optional ByVal PRICE_CATGY_CODE As String = "")
        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Items")
        Application.DoEvents()

        Try

            Dim SQL As String = "Select * from " & ICTINVA1 & " ICTINVA1"

            If COL = "" Then
                SQL = SQL & " where PRICE_CATGY_CODE = '" & PRICE_CATGY_CODE & "'"
                Dim PRICE_CATGY_DESC As String = dst.Tables("ICTPCAT1").Rows.Find(PRICE_CATGY_CODE).Item("PRICE_CATGY_DESC") & ""

                grdICTINVA1.Text = grdICTINVAT.Text & ", Price Category " & PRICE_CATGY_CODE & " (" & PRICE_CATGY_DESC & ")"
            Else

                grdICTINVA1.Text = grdICTINVA0.Text

                Dim TC As String = ""
                For Each T As String In IOT.Keys
                    If IOT(T) = Mid(COL, 3, 3) Then
                        TC = T
                        grdICTINVA1.Text = grdICTINVA0.Text & ", " & Mid(COL, 3, 3) & " Items Only"
                        Exit For
                    End If
                Next

                Dim C As String = "="
                Dim CZ As String = ""

                'Dim tlbopt As UltraWinToolbars.OptionSet = tlb.OptionSets("OnlyOneUp")

                'Select Case tlbopt.SelectedTool.Key
                '    Case "Show All Items in Selected Cell Only"
                '        C = "="
                '    Case "All Items in and below Selected Cell"
                '        C = ">="
                '        CZ = "(At Least) "
                '    Case "All Items in and above Selected Cell"
                '        C = "<="
                '        CZ = "(At Most) "
                'End Select

                Select Case optCell.Value
                    Case "0"
                        C = "="
                    Case "1"
                        C = ">="
                        CZ = "(At Least) "
                    Case "2"
                        C = "<="
                        CZ = "(At Most) "
                End Select

                'If MOS = 0 Then ' if you double-click the 0 line, you get only 0's
                '    C = "="
                '    CZ = ""
                'End If

                SQL &= " where NVL(ICTINVA1.MOS,0) " & C & " " & CStr(MOS)

                If MOS = 0 Then
                    'SQL &= " where MOS is Null"
                    grdICTINVA1.Text = grdICTINVA1.Text & ", No Activity Last " & CStr(MOSMAX) & " " & optWM.Text
                Else
                    'SQL &= " where MOS = " & CStr(MOS)
                    grdICTINVA1.Text = grdICTINVA1.Text & ", Active in " & CZ & CStr(MOS) & " of the Last " & CStr(MOSMAX) & " " & optWM.Text
                End If

                If Mid(COL, 3, 3) <> "ALL" Then
                    SQL &= " and ICTINVA1.ITEM_ORDER_CODE = '" & TC & "'"
                End If

                SQL = SQL & sqlw
            End If

            Call Fill_Records("ICTINVA1", "", , SQL)

            SplitContainer6.Panel2Collapsed = True
            tabMain.Tabs("Items").Enabled = True
            tabMain.SelectedTab = tabMain.Tabs("Items")

        Catch ex As Exception

            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Please Alert ABS that you received this error")

        End Try

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()

    End Sub

    Private Sub grpShowTrend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpShowTrend.Click

    End Sub

    Private Sub optTrendUnit_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTrendUnit.ValueChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_Trend()
    End Sub

    Private Sub tabSummary_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSummary.SelectedTabChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Setup_tabSummary()
    End Sub

    Private Sub optTrendCode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTrendCode.ValueChanged
        If Me.Visible = False Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        If optTrendCode.Value = "C" And optPROD_CATGY_CODE.Value = "S" Then
            optPROD_CATGY_CODE.Value = "A"
        ElseIf optTrendCode.Value = "V" And optVEND_CODE.Value = "S" Then
            optVEND_CODE.Value = "A"
        End If
        Call Fetch_Trend()
    End Sub

    Sub Fetch_Trend()

        Dim chtIsVisible As Boolean = chtICTINVAT.Visible
        chtICTINVAT.Visible = False

        chtICTINVAT.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Fetching Trend Data")
        Application.DoEvents()


        Me.SuspendLayout()
        'grdICTINVAT.SuspendLayout()
        'grdICTINVAT.EventManager.AllEventsEnabled = False
        'chtICTINVAT.SuspendLayout()


        dst.Tables("ICTINVAT").Rows.Clear()

        Dim sql As String = ""
        Dim S(1) As String
        Dim CAPTIONS(1) As String
        Dim TABLE_NAME As String = ""
        Dim tt As String = ""
        Select Case optTrendCode.Value
            Case "C"
                TABLE_NAME = "ICTCATG1"
                S = New String() {"PROD_CATGY_CODE", "PROD_CATGY_DESC"}
                CAPTIONS(0) = "Category"
                CAPTIONS(1) = "Description"
                tt = "All Vendors in this Product Category"

                If selectedProdCatgyCode.Length > 0 Then
                    viewICTINVAT.RowFilter = "CODE_VALUE = '" & selectedProdCatgyCode & "'"
                Else
                    viewICTINVAT.RowFilter = String.Empty
                End If

            Case "V"
                TABLE_NAME = "APTVEND1"
                S = New String() {"VEND_CODE", "VEND_NAME"}
                CAPTIONS(0) = "Vendor"
                CAPTIONS(1) = "Name"
                tt = "All Price Categories in this Vendor"

                If selectedVendor.Length > 0 Then
                    viewICTINVAT.RowFilter = "CODE_VALUE = '" & selectedVendor & "'"
                Else
                    viewICTINVAT.RowFilter = String.Empty
                End If

            Case "P"
                TABLE_NAME = "ICTPCAT1"
                S = New String() {"PRICE_CATGY_CODE", "PRICE_CATGY_DESC"}
                CAPTIONS(0) = "Price"
                CAPTIONS(1) = "Description"
                tt = "All Items in this Price Category"

                If selectedPriceCatgyCode.Length > 0 Then
                    viewICTINVAT.RowFilter = "CODE_VALUE = '" & selectedPriceCatgyCode & "'"
                Else
                    viewICTINVAT.RowFilter = String.Empty
                End If

        End Select

        grdICTINVAT.DataSource = viewICTINVAT

        With grdICTINVAT.DisplayLayout.Bands(0)
            .Columns("CODE_VALUE").Header.Caption = CAPTIONS(0) ' ASCMAIN1.Make_Caption(S(0))
            .Columns("DESC_VALUE").Header.Caption = CAPTIONS(1) ' ASCMAIN1.Make_Caption(S(1))
        End With

        Dim RL() As String
        Dim CL() As String
        ReDim CL(MOSMAX)

        Dim U As String = optTrendUnit.Value

        sql = "Select " _
        & TABLE_NAME & "." & S(0) & " CODE_VALUE, " _
        & TABLE_NAME & "." & S(1) & " DESC_VALUE"
        For i As Integer = 0 To MOSMAX
            sql &= ", SUM (" & U & "" & Format(i, "00") & ") " & U & Format(i, "00")
        Next
        sql &= " from " & ICTINVA1 & " ICTINVA1, " & TABLE_NAME _
        & " where " & TABLE_NAME & "." & S(0) & " (+) = ICTINVA1." & S(0) _
        & " AND " & TABLE_NAME & "." & S(0) & " IS NOT NULL " _
        & sqlw _
        & " group by " _
        & TABLE_NAME & "." & S(0) & ", " _
        & TABLE_NAME & "." & S(1)

        grdICTINVAT.DisplayLayout.Bands(0).Columns("CODE_VALUE").Width = 100
        grdICTINVAT.DisplayLayout.Bands(0).Columns("DESC_VALUE").Width = 200

        With dst.Tables("ICTINVAT").Columns
            If .Count > 2 Then
                For i As Integer = .Count - 1 To 2 Step -1
                    Dim X_COLUMN_NAME As String = .Item(i).ColumnName
                    .Remove(X_COLUMN_NAME)
                    With grdICTINVAT.DisplayLayout.Bands(0)
                        .Summaries.Remove(.Summaries(X_COLUMN_NAME))
                    End With
                Next

            End If

            'grdICTINVAT.DataSource = dst.Tables("ICTINVAT")

            For i As Integer = 0 To MOSMAX
                Dim X_COLUMN_NAME As String = U & Format(i, "00")
                'If U = "S" Then
                '    .Add(X_COLUMN_NAME, GetType(System.Decimal))
                'Else
                '    .Add(X_COLUMN_NAME, GetType(System.Decimal))
                'End If
                If U = "S" Then
                    .Add(X_COLUMN_NAME, GetType(System.Double))
                Else
                    .Add(X_COLUMN_NAME, GetType(System.Int32))
                End If

                With grdICTINVAT.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME)
                    '.Width = grdICTINVA1.DisplayLayout.Bands(0).Columns(U & "00").Width * 1.5
                    .Width = grdICTINVAT.DisplayLayout.Bands(0).Columns("CODE_VALUE").Width
                    '.Header.Caption = row(0)
                    '.CellAppearance.BackColor = grdICTINVA1.DisplayLayout.Bands(0).Columns(slq & "00").CellAppearance.BackColor
                    .Format = "#,##0" ' grdICTINVA1.DisplayLayout.Bands(0).Columns(SQ & "00").Format

                    If U = "S" Then
                        .CellAppearance.BackColor = Color.LightGreen
                    Else
                        .CellAppearance.BackColor = Color.AliceBlue
                    End If

                    .Header.Appearance.TextHAlign = HAlign.Right
                    .CellAppearance.TextHAlign = HAlign.Right
                    .Hidden = False
                End With

                Call Create_Summary(grdICTINVAT, U & Format(i, "00"))
            Next

            For i As Integer = 0 To MOSMAX
                Dim X_COLUMN_NAME As String = U & Format(i, "00")
                grdICTINVAT.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME).MaskInput = "#,##0"
                If i <> 0 Then
                    CL(i - 1) = grdICTINVA1.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME).Header.Caption
                End If
                'dst.Tables("ICTINVAT").Columns(X_COLUMN_NAME).Caption = grdICTINVAT.DisplayLayout.Bands(0).Columns(X_COLUMN_NAME).Header.Caption
            Next

        End With


        Dim OTHERS() As Decimal
        ReDim OTHERS(MOSMAX)

        Dim top As New List(Of String)
        If Not chkTop.Checked Then
            Try
                Call Fill_Records("ICTINVAT", "", True, sql)
            Catch ex As Exception
                Call Fill_Records("ICTINVAT", "", True, sql)
            End Try
            'Stop
        Else
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)

            If chkTop.Checked Then
                Dim rc As Integer = 0
                For Each row As DataRow In tbl.Select("", U & "00 DESC")
                    rc = rc + 1
                    If rc > Val(numTOP.Value) Then
                        Exit For
                    Else
                        top.Add(row("CODE_VALUE") & "")
                    End If
                Next
            End If


            With dst.Tables("ICTINVAT")
                For Each row As DataRow In tbl.Rows
                    If Not chkTop.Checked Or top.Contains(row("CODE_VALUE") & "") Then
                        Dim rowICTINVAT As DataRow = .NewRow
                        rowICTINVAT.ItemArray = row.ItemArray
                        .Rows.Add(rowICTINVAT)
                    Else
                        For i As Integer = 0 To MOSMAX
                            OTHERS(i) += Val(row(U & Format(i, "00")) & "")
                        Next
                    End If
                Next
            End With

            If chkTop.Checked Then
                Dim row As DataRow = dst.Tables("ICTINVAT").NewRow
                row("CODE_VALUE") = "*"
                row("DESC_VALUE") = "All Others"
                For i As Integer = 0 To MOSMAX
                    row(U & Format(i, "00")) = OTHERS(i)
                Next
                dst.Tables("ICTINVAT").Rows.Add(row)
            End If

        End If

        'dst.Tables("ICTINVAT").Merge(tbl, False)
        'Call Fill_Records("ICTINVAT", "", True, sql)

        ' LATER TURN THIS IN TO A CALLABLE ROUTINE:
        With grdICTINVAT.DisplayLayout
            .Bands("ICTINVAT").SortedColumns.Clear()
            .Bands("ICTINVAT").SortedColumns.Add("CODE_VALUE", False)
            ' not nec if we set the sort on an empty grid and then fill the data
            ' rows(0) refers to the 1st visible row
        End With
        If grdICTINVAT.Rows.Count <> 0 Then
            grdICTINVAT.ActiveRow = grdICTINVAT.Rows(0)
        End If

        chtICTINVAT.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom

        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtICTINVAT.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtICTINVAT.LabelHash = labelHash

        chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtICTINVAT.Tooltips.FormatString = "<HIGHLOW>"


        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("ICTINVAT").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("ICTINVAT").Select("", "CODE_VALUE")
            RL(RLi) = row("CODE_VALUE") & ":" & row("DESC_VALUE")
            RLi += 1
        Next
        chtICTINVAT.Data.SetRowLabels(RL)
        chtICTINVAT.Data.SetColumnLabels(CL)

        'chtICTINVAT.DataSource = dst.Tables("ICTINVAT")
        chtICTINVAT.DataSource = New DataView(dst.Tables("ICTINVAT"), "", "CODE_VALUE", DataViewRowState.CurrentRows)
        chtICTINVAT.Data.IncludeColumn(2, False)

        chtICTINVAT.DataBind()

        chtICTINVAT.Visible = chtIsVisible


        'chtICTINVAT.ResumeLayout()

        With grdICTINVAT.DisplayLayout.Bands(0)
            For i As Integer = 0 To MOSMAX
                .Columns(U & Format(i, "00")).Header.Caption = grdICTINVA1.DisplayLayout.Bands(0).Columns(U & Format(i, "00")).Header.Caption
                ' .Columns(U & Format(i, "00")).Format = grdICTINVA1.DisplayLayout.Bands(0).Columns(U & Format(i, "00")).Format
            Next
        End With

        'grdICTINVAT.ResumeLayout()
        'grdICTINVAT.EventManager.AllEventsEnabled = True

        grdICTINVAT.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals; Double-Click a row to Drill down to " & tt

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()

    End Sub

    Private Sub SplitContainer7_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer7.SplitterMoved

    End Sub

    Private Sub chkChart_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkChart.CheckedChanged
        If Not Me.Visible Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        SplitContainer7.Panel2Collapsed = Not chkChart.Checked
        chtICTINVAT.Visible = chkChart.Checked
    End Sub

    Private Sub chtICTINVAT_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles chtICTINVAT.MouseHover
        'chtICTINVAT.Tooltips.FormatString = "tooltip"
    End Sub

    Private Sub optChartD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optChartD.ValueChanged
        trkbrXAxis.Visible = (optChartD.Value = "2")
        trkbrYAxis.Visible = (optChartD.Value = "2")
        lblXAxis.Visible = (optChartD.Value = "2")
        lblYAxis.Visible = (optChartD.Value = "2")

        If optChartD.Value = "2" Then
            chtICTINVAT.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.LineChart
        Else
            chtICTINVAT.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.LineChart3D
        End If
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtICTINVAT.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0

    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtICTINVAT.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0

    End Sub

    Private Sub cmdChartPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChartPrint.Click
        printDialog1 = New System.Windows.Forms.PrintDialog
        printDialog1.Document = chtICTINVAT.PrintDocument
        If Me.printDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
        End If
    End Sub

    Private Sub comboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles comboBox1.SelectedIndexChanged
        chtICTINVAT.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(ColorModels), Me.comboBox1.SelectedItem.ToString()), ColorModels)
    End Sub

    Private Sub chkTop_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTop.CheckedChanged
        If Not Me.Visible Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_Trend()
    End Sub

    Private Sub numTOP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numTOP.ValueChanged
        If Not Me.Visible Then Exit Sub
        If dst.Tables.Count = 0 Then Exit Sub
        Call Fetch_Trend()
    End Sub

    Private Sub grdICTINVAT_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTINVAT.DoubleClickRow
        Try
            With grdICTINVAT
                If .ActiveRow Is Nothing Then
                    Exit Sub
                End If

                Select Case optTrendCode.Value
                    Case "C"
                        grdICTCATG1.Selected.Rows.Clear()
                        grdICTCATG1.ActiveRow = grdICTCATG1.Rows.GetRowWithListIndex( _
                        dst.Tables("ICTCATG1").Rows.IndexOf(dst.Tables("ICTCATG1").Rows _
                        .Find(grdICTINVAT.ActiveRow.Cells("CODE_VALUE").Text)))
                        grdICTCATG1.Selected.Rows.Add(grdICTCATG1.ActiveRow)
                        optPROD_CATGY_CODE.Value = "S"
                        optTrendCode.Value = "V"
                    Case "V"
                        grdAPTVEND1.Selected.Rows.Clear()
                        grdAPTVEND1.ActiveRow = grdAPTVEND1.Rows.GetRowWithListIndex( _
                        dst.Tables("APTVEND1").Rows.IndexOf(dst.Tables("APTVEND1").Rows _
                        .Find(grdICTINVAT.ActiveRow.Cells("CODE_VALUE").Text)))
                        grdAPTVEND1.Selected.Rows.Add(grdAPTVEND1.ActiveRow)
                        optVEND_CODE.Value = "S"
                        optTrendCode.Value = "P"
                    Case "P"
                        Dim PRICE_CATGY_CODE As String = grdICTINVAT.ActiveRow.Cells("CODE_VALUE").Text
                        Setup_ICTINVA1("", 0, PRICE_CATGY_CODE)
                End Select

                grdICTINVAT.DisplayLayout.Override.RowSelectors = DefaultableBoolean.True
                grdICTINVAT.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

            End With

        Catch ex As Exception
            Dim zMsg As String = "The following error was trapped: " & vbCrLf
            zMsg &= ex.Message & vbCrLf & vbCrLf
            zMsg &= "Click OK to resume what you were doing."
            MessageBox.Show(zMsg, "Trap Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub chkShowFilters_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowFilters.CheckedChanged
        Call Show_Filter(grdICTINVA1, chkShowFilters.Checked)
    End Sub

End Class


Public Class MyCustomTooltip
    Implements IRenderLabel

    Public Sub New()

    End Sub 'New

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

    End Function 'ToString 
End Class 'MyCustomTooltip
