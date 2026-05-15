Imports System.Math
Imports System.Drawing

Public Class APRACRX1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Shadows SUBT As String = ""

    Dim APTACRX1 As String
    Dim sqlAPTACRX1 As String = ""

    Dim APTACRCX As String = String.Empty
    Dim sqlAPTACRCX As String

    Dim APTACRX0 As String
    Dim sqlAPTACRX0 As String = ""

    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid
    Dim grdASTEXPT3 As New UltraWinGrid.UltraGrid

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        Get_PARM("APTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        SUBT = ""

        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Receipts & Pre-Payments Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Receipts & Pre-Payments Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "APTACRC1.CTL_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"

        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Receipts & Pre-Payments Posted in " & xRYP0_legend
            Else
                SUBT = "Receipts & Pre-Payments Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "APTACRC1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"

        ElseIf Absx1.optFor("RANGE").Value = "A" Then
            SUBT = "Selected Matches"
            RWU = "N"
        End If

        If sqlw <> "" Then sqlw &= " and "
        sqlw &= "APTACRC1.ACCRUAL_CODE = 'TRF' and (APTACRC1.CTL_STATUS = '0' or (APTACRC1.CTL_STATUS = '1' and APTACRC1.PPD_IND = '1' and NVL(APTACRC1.PPD_MATCHED,'0') = '0'))"

        If ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If

        'sqlw &= SQLA_filter("WHSE_CODE", "ICTIREC1")
        'sqlw &= SQLA_filter("VEND_CODE", "ICTIREC1")
        'sqlw &= SQLA_filter("PO_ORDER_NO", "ICTIREC1")
        'sqlw &= SQLA_filter("RECEIPT_NO", "ICTIREC1")

        Prepare_dst(True, sqlw)

        Check_if_Empty("APTACRX1")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = "Matched using BOL"
        Generate_Report(RPT, , SUBT)

        RPT = "APRACRX2"
        SUBT = "Matched but Variance exceeds Tolerance"
        Generate_Report(RPT, , SUBT)

        Print_GL()

        Prepare_Data_Extracts()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "A" Then
                If tblASTDSQLA.Select("CODE_VALUES <> ''").Length = 0 Then
                    EMsg &= vbCr & "You must Specify some Filter Criteria"
                End If
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

        ' Dim VOUCHER_NO As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")

        ASCMAIN1.sql = $"Update APTACRC1 SET PPD_MATCHED = '1', PPD_MATCHED_XNO = '{XNO}'" & vbCrLf _
            & " where CTL_NO IN (" & vbCrLf _
            & $"Select CTL_NO from {APTACRX1} where BOL_NO IN (" & vbCrLf _
            & $"Select BOL_NO from {APTACRX0} where MATCH_FAIL = 'MATCH')" & vbCrLf _
            & " and PPD_IND = '1'" & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Update APTACRC1 SET CTL_STATUS = '1', PPD_MATCHED_XNO = '{XNO}'" & vbCrLf _
            & $", BOL_NO_MATCHED = (Select BOL_NO from {APTACRX1} where CTL_NO = APTACRC1.CTL_NO)" & vbCrLf _
            & " where CTL_NO IN (" & vbCrLf _
            & $"Select CTL_NO from {APTACRX1} where BOL_NO IN (" & vbCrLf _
            & $"Select BOL_NO from {APTACRX0} where MATCH_FAIL = 'MATCH')" & vbCrLf _
            & " and NVL(PPD_IND,'0') = '0'" & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = $"Update APTACRC1 X Set OPS_YYYYPP_MATCHED = '{ASCMAIN1.CYP}', VOUCHER_NO_MATCHED = (
        '    Select VOUCHER_NO from APTACRC1 
        '    where PPD_MATCHED_XNO = '{XNO}' AND (NVL(PPD_IND,'0') = '1' AND SOURCE_DOC_NO = X.BOL_NO_MATCHED)
        '    )
        '    where PPD_MATCHED_XNO = '{XNO}' AND (NVL(PPD_IND,'0') = '0')"
        ASCMAIN1.sql = $"Update APTACRC1 X Set OPS_YYYYPP_MATCHED = '{ASCMAIN1.CYP}'
            where PPD_MATCHED_XNO = '{XNO}' AND (NVL(PPD_IND,'0') = '0')"
        ASCDATA1.ExecuteSQL()

        For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select("COST_VAR_ITEM <> 0")
            Dim CTL_NO As String = rowAPTACRX1.Item("CTL_NO")
            Dim ITEM_CODE As String = rowAPTACRX1.Item("ITEM_CODE")
            Dim COST_VAR_ITEM As Decimal = Val(rowAPTACRX1.Item("COST_VAR_ITEM") & "")
            Dim TPV_ADJ As Decimal = Val(rowAPTACRX1.Item("TPV_ADJ") & "")
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & $" Update ICTIVAR1 Set TV_EXP = NVL(TV_EXP,0) + {CStr(COST_VAR_ITEM)}" & vbCrLf _
                & $" where ITEM_CODE = '{ITEM_CODE}' and OPS_YYYYPP = '{ASCMAIN1.CYP}';" & vbCrLf _
                & "  If SQL%NOTFOUND Then" & vbCrLf _
                & "   Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, TV_EXP)" & vbCrLf _
                & $"    Values ('{ITEM_CODE}', '{ASCMAIN1.CYP}', {CStr(COST_VAR_ITEM)});" & vbCrLf _
                & "  End If;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            Debug.Print(CTL_NO & ":" & ITEM_CODE & ":" & CStr(COST_VAR_ITEM))

            ASCMAIN1.sql = $"Update APTACRC1 Set COST_VAR_ITEM = {CStr(COST_VAR_ITEM)} where CTL_NO = '{CTL_NO}'"
            'ASCMAIN1.sql = $"Update APTACRC1 Set COST_VAR_ITEM = {CStr(COST_VAR_ITEM)}, TPV_ADJ = {CStr(TPV_ADJ)} where CTL_NO = '{CTL_NO}'"
            ASCDATA1.ExecuteSQL()
        Next

        For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select("MATCH_FAIL = 'MATCH'")
            Dim CTL_NO As String = rowAPTACRX1.Item("CTL_NO")
            Dim BOL_NO As String = rowAPTACRX1.Item("BOL_NO")
            Dim CTL_NO_MATCHED As String = rowAPTACRX1.Item("CTL_NO_MATCHED")

            ASCMAIN1.sql = $"Update APTACRC1 Set CTL_NO_MATCHED = '{CTL_NO_MATCHED}', BOL_NO = '{BOL_NO}' where CTL_NO = '{CTL_NO}'"
            ASCDATA1.ExecuteSQL()
        Next

        '  Stop ' NOT SURE ABOUT THIS - THERE IS THE ORIG, THE REVERSAL, AND NEW CTL - AND IT IS WIERD THAT THE NEW CTL IS NOT COMING OUT IN THE REPORT
        For Each rowAPTACRC0 As DataRow In dst.Tables("APTACRC0").Select("ISNULL(CTL_NO_MATCHED_PREV,'?') <> '?'")
            Dim CTL_NO_MATCHED As String = rowAPTACRC0.Item("CTL_NO_MATCHED")
            Dim CTL_NO_MATCHED_PREV As String = rowAPTACRC0.Item("CTL_NO_MATCHED_PREV")

            ASCMAIN1.sql = $"Update APTACRC0 Set CTL_NO_MATCHED_NEXT = '{CTL_NO_MATCHED}' where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            ASCDATA1.ExecuteSQL()
        Next


        Update_Record_TDA("APTACRC0")

        If MENU_ITEM_OBJECT = "APRACRX1" Then
            GL_Update()
        End If

    End Sub


    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        ASCMAIN1.sql = "Select APTACRX1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from " & APTACRX1 & " APTACRX1, ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = APTACRX1.ITEM_CODE" & vbCrLf _
            & "   And MATCH_FAIL = 'MATCH'"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        tbl.Columns.Add("COST_VAR", GetType(System.Decimal), "ISNULL(COST_ACT,0) - ISNULL(COST_ACC,0)")

        grdASTEXPT1.DataSource = tbl

        grdASTEXPT1.Text = "PPD Tariff Matches within Tolerance"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "BOL_NO", "BOL No", 120, , "Count", Color.Violet)
        Set_DX_Column(grdASTEXPT1, "SOURCE_DOC_NO", "Source Doc No", 160, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE_ACC", "Vendor", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT1, "ACCRUAL_CODE", "Type", 50, , , Color.CornflowerBlue)

        Set_DX_Column(grdASTEXPT1, "CTL_NO", "Ctl No", 100, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "CTL_DATE", "Date", 100, "MM/dd/yy", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "CTL_NOTE", "Note", 200, , , Color.LightPink)

        Set_DX_Column(grdASTEXPT1, "COST_ACC", "Accrued", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "TPV_ADJ", "TPV Adj", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "COST_ACT", "Pre-Paid", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "COST_VAR", "Variance", 120, "#,##0.00", "Sum", Color.LightGreen)

        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "YP", 75, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VOUCHER_NO", "Voucher No", 100, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_NO", "Receipt No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_LNO", "Ln", 30, "##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_NO", "PO No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_LNO", "Ln", 30, "##0", , Color.LightBlue)

        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130, , , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "QTY_REC", "Qty Rec", 90, "#,##0", "Sum", Color.LightBlue)

        'Sort_grdColumns(grdASTEXPT1, "SOURCE_DOC_NO")

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT1, "BOL_NO,CTL_NO")





        With grdASTEXPT2 ' this grid is instantiated in code up in form declarations
            grdASTEXPT2.Name = "grdASTEXPT2"
            If Not GRDs.ContainsKey("ASTEXPT2") Then

                ' this SECTION should be placed in ASCMAIN1
                tabDataExports.Tabs.Add()

                GRDs.Add(Mid(.Name, 4), grdASTEXPT2)
                Add_Handlers_grd(grdASTEXPT2)

                .DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy

                .Parent = tabDataExports.Tabs(1).TabPage
                .Text = "Grid Caption set below"

                .Dock = System.Windows.Forms.DockStyle.Fill
                'tabDataExports.Tabs(1).Text = .Text

                .DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
                .DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

                tabDataExports.Tabs(1).Text = .Text
                .DisplayLayout.Override.AllowGroupBy = DefaultableBoolean.True
                .DisplayLayout.GroupByBox.Hidden = False
                .DisplayLayout.MaxColScrollRegions = 1
                .DisplayLayout.MaxRowScrollRegions = 1

                '.DisplayLayout.Override.RowAppearance.BorderColor = System.Drawing.Color.Silver
                .DisplayLayout.Override.RowAppearance = grdASTEXPT1.DisplayLayout.Override.RowAppearance
            End If

            ASCMAIN1.grdInitializeLayout(grdASTEXPT2)
        End With


        grdASTEXPT2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        ASCMAIN1.sql = "Select APTACRX1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from " & APTACRX1 & " APTACRX1, ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = APTACRX1.ITEM_CODE" & vbCrLf _
            & "   and MATCH_FAIL = 'FAIL'"
        Dim tbl2 As DataTable = ASCDATA1.GetDataTable
        tbl2.Columns.Add("COST_VAR", GetType(System.Decimal), "ISNULL(COST_ACT,0) - ISNULL(COST_ACC,0)")

        grdASTEXPT2.DataSource = tbl2

        grdASTEXPT2.Text = "PPD Tariff Matches exceeding Tolerance"
        ' UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(1).Visible = True
        tabDataExports.Tabs(1).Text = grdASTEXPT2.Text

        Set_DX_Column(grdASTEXPT2, "")

        ' ASCMAIN1.grdInitializeLayout(grdASTEXPT2)


        Set_DX_Column(grdASTEXPT2, "BOL_NO", "BOL No", 120, , "Count", Color.Violet)
        Set_DX_Column(grdASTEXPT2, "SOURCE_DOC_NO", "Source Doc No", 160, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "VEND_CODE_ACC", "Vendor", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "ACCRUAL_CODE", "Type", 50, , , Color.CornflowerBlue)

        Set_DX_Column(grdASTEXPT2, "CTL_NO", "Ctl No", 100, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT2, "CTL_DATE", "Date", 100, "MM/dd/yy", , Color.LightPink)
        Set_DX_Column(grdASTEXPT2, "CTL_NOTE", "Note", 200, , , Color.LightPink)

        Set_DX_Column(grdASTEXPT2, "COST_ACC", "Accrued", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "TPV_ADJ", "TPV Adj", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "COST_ACT", "Pre-Paid", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "COST_VAR", "Variance", 120, "#,##0.00", "Sum", Color.LightGreen)

        Set_DX_Column(grdASTEXPT2, "OPS_YYYYPP", "YP", 75, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "VOUCHER_NO", "Voucher No", 100, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "RECEIPT_NO", "Receipt No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "RECEIPT_LNO", "Ln", 30, "##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "PO_ORDER_NO", "PO No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "PO_ORDER_LNO", "Ln", 30, "##0", , Color.LightBlue)

        Set_DX_Column(grdASTEXPT2, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        Set_DX_Column(grdASTEXPT2, "ITEM_DESC", "Description", 130, , , Color.Gold)
        Set_DX_Column(grdASTEXPT2, "QTY_REC", "Qty Rec", 90, "#,##0", "Sum", Color.LightBlue)

        'Sort_grdColumns(grdASTEXPT2, "SOURCE_DOC_NO")

        grdASTEXPT2.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT2, "BOL_NO,CTL_NO")






        With grdASTEXPT3 ' this grid is instantiated in code up in form declarations
            grdASTEXPT3.Name = "grdASTEXPT3"
            If Not GRDs.ContainsKey("ASTEXPT3") Then

                ' this SECTION should be placed in ASCMAIN1
                tabDataExports.Tabs.Add()

                GRDs.Add(Mid(.Name, 4), grdASTEXPT3)
                Add_Handlers_grd(grdASTEXPT3)

                .DisplayLayout.ViewStyleBand = UltraWinGrid.ViewStyleBand.OutlookGroupBy

                .Parent = tabDataExports.Tabs(2).TabPage
                .Text = "Grid Caption set below"

                .Dock = System.Windows.Forms.DockStyle.Fill
                'tabDataExports.Tabs(2).Text = .Text

                .DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
                .DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement

                tabDataExports.Tabs(2).Text = .Text
                .DisplayLayout.Override.AllowGroupBy = DefaultableBoolean.True
                .DisplayLayout.GroupByBox.Hidden = False
                .DisplayLayout.MaxColScrollRegions = 1
                .DisplayLayout.MaxRowScrollRegions = 1

                '.DisplayLayout.Override.RowAppearance.BorderColor = System.Drawing.Color.Silver
                .DisplayLayout.Override.RowAppearance = grdASTEXPT1.DisplayLayout.Override.RowAppearance
            End If

            ASCMAIN1.grdInitializeLayout(grdASTEXPT3)
        End With


        grdASTEXPT3.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        ASCMAIN1.sql = "Select APTACRX1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from " & APTACRX1 & " APTACRX1, ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = APTACRX1.ITEM_CODE" & vbCrLf _
            & "   and MATCH_FAIL = 'MATCH' and COST_VAR_ITEM <> 0"
        Dim tbl3 As DataTable = ASCDATA1.GetDataTable
        tbl3.Columns.Add("COST_VAR", GetType(System.Decimal), "ISNULL(COST_ACT,0) - ISNULL(COST_ACC,0)")

        grdASTEXPT3.DataSource = tbl3

        'grdASTEXPT3.DataSource = dst.Tables("APTACRX1").Select("MATCH_FAIL = 'MATCH' and COST_VAR_ITEM <> 0")

        grdASTEXPT3.Text = "TPV Summary by Item"
        ' UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(2).Visible = True
        tabDataExports.Tabs(2).Text = grdASTEXPT3.Text

        Set_DX_Column(grdASTEXPT3, "")

        'ASCMAIN1.grdInitializeLayout(grdASTEXPT3)

        grdASTEXPT3.Visible = True

        Set_DX_Column(grdASTEXPT3, "BOL_NO", "BOL No", 120, , "Count", Color.Violet)
        Set_DX_Column(grdASTEXPT3, "SOURCE_DOC_NO", "Source Doc No", 160, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT3, "VEND_CODE_ACC", "Vendor", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT3, "ACCRUAL_CODE", "Type", 50, , , Color.CornflowerBlue)

        Set_DX_Column(grdASTEXPT3, "CTL_NO", "Ctl No", 100, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT3, "CTL_DATE", "Date", 100, "MM/dd/yy", , Color.LightPink)
        Set_DX_Column(grdASTEXPT3, "CTL_NOTE", "Note", 200, , , Color.LightPink)

        Set_DX_Column(grdASTEXPT3, "COST_ACC", "Accrued", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT3, "TPV_ADJ", "TPV Adj", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT3, "COST_VAR_ITEM", "TPV", 120, "#,##0.00", "Sum", Color.LimeGreen)
        Set_DX_Column(grdASTEXPT3, "COST_ACT", "Pre-Paid", 120, "#,##0.00", "Sum", Color.LightGreen)
        Set_DX_Column(grdASTEXPT3, "COST_VAR", "Variance", 120, "#,##0.00", "Sum", Color.LightGreen)

        Set_DX_Column(grdASTEXPT3, "OPS_YYYYPP", "YP", 75, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT3, "VOUCHER_NO", "Voucher No", 100, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT3, "RECEIPT_NO", "Receipt No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT3, "RECEIPT_LNO", "Ln", 30, "##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT3, "PO_ORDER_NO", "PO No", 70, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT3, "PO_ORDER_LNO", "Ln", 30, "##0", , Color.LightBlue)

        Set_DX_Column(grdASTEXPT3, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        Set_DX_Column(grdASTEXPT3, "ITEM_DESC", "Description", 130, , , Color.Gold)
        Set_DX_Column(grdASTEXPT3, "QTY_REC", "Qty Rec", 90, "#,##0", "Sum", Color.LightBlue)

        'Sort_grdColumns(grdASTEXPT3, "SOURCE_DOC_NO")

        grdASTEXPT3.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT3, "BOL_NO,CTL_NO")

    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        sqlAPTACRX1 = "Select APTACRC1.*, ICTITEM1.ITEM_DESC, ICTIREC2.QTY_REC" & vbCrLf _
            & " from APTACRC1, ICTITEM1, ICTIREC2" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_NO (+) = APTACRC1.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO (+) = APTACRC1.RECEIPT_LNO"

        ASCMAIN1.sql = sqlAPTACRX1 & " and ROWNUM < 1"
        APTACRX1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add Primary Key (CTL_NO)")
        'ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add BOL_NO VARCHAR2(20)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add MATCH_FAIL VARCHAR2(20)")

        'ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add DEF_TOLERANCE NUMBER(6,2)")
        'ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add COST_ACC_TOTAL NUMBER(13,2)")
        'ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add TPV_ADJ_TOTAL NUMBER(13,2)")
        'ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add LNO_COUNT NUMBER(5,0)")



        ASCMAIN1.sql = "Select APTACRX1.* " & vbCrLf _
            & " from " & APTACRX1 & " APTACRX1"
        Create_TDA(dst.Tables.Add("APTACRX1"), APTACRX1, "**", 0, True,, 1) ', "COST_VAR_ITEM")
        dst.Tables("APTACRX1").Columns.Add("COST_VAR", GetType(System.Decimal), "ISNULL(COST_ACT,0) - ISNULL(COST_ACC,0) - ISNULL(TPV_ADJ,0)")


        sqlAPTACRCX = "Select APTACRC0.CTL_NO_MATCHED" & vbCrLf _
                & " from APTACRC0"
        ASCMAIN1.sql = sqlAPTACRCX & " where ROWNUM < 1"
        APTACRCX = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add Primary Key (CTL_NO_MATCHED)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add DEF_TOLERANCE NUMBER(6,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add COST_ACT_TOTAL NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add COST_ACC_TOTAL NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add TPV_ADJ_TOTAL NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add CALC_VARIANCE NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add LNO_COUNT NUMBER(5,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add SOURCE_DOC_NO VARCHAR2(20)")


        'sqlAPTACRX0 = "Select Distinct Z.*" & vbCrLf _
        '    & ", Case when VAR_TOLERANCE >= ABS(COST_VAR) THEN 'MATCH' else 'FAIL' END MATCH_FAIL" & vbCrLf _
        '    & " from (" & vbCrLf _
        '    & "Select X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR, APTACRX1.VAR_TOLERANCE" & vbCrLf _
        '    & $" from {APTACRX1} APTACRX1, (" & vbCrLf _
        '    & "Select BOL_NO" & vbCrLf _
        '    & ", SUM (NVL(COST_ACT,0)) COST_ACT, SUM (NVL(COST_ACC,0)) COST_ACC, SUM (NVL(TPV_ADJ,0)) TPV_ADJ, SUM (CASE WHEN NVL(PPD_IND,'0') = '0' THEN 1 ELSE 0 END) LINE_ITEMS" & vbCrLf _
        '    & $" from {APTACRX1} APTACRX1  " & vbCrLf _
        '    & " group by BOL_NO) X" & vbCrLf _
        '    & " where APTACRX1.BOL_NO = X.BOL_NO and APTACRX1.PPD_IND = '1'" & vbCrLf _
        '    & ") Z"



        '        Select Case X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR, X.VAR_TOLERANCE
        ' from(SELECT DISTINCT BOL_NO FROM ASW74265 WHERE PPD_IND = '1') APTACRX1, (
        'Select Case BOL_NO
        ', SUM (NVL(COST_ACT,0)) COST_ACT, SUM (NVL(COST_ACC,0)) COST_ACC, SUM (NVL(TPV_ADJ,0)) TPV_ADJ, SUM (CASE WHEN NVL(PPD_IND,'0') = '0' THEN 1 ELSE 0 END) LINE_ITEMS
        ', MAX (VAR_TOLERANCE) VAR_TOLERANCE
        ' From ASW74265 APTACRX1
        ' Group By BOL_NO) X
        ' Where APTACRX1.BOL_NO = X.BOL_NO
        'And APTACRX1.BOL_NO = 'SA01181675'

        sqlAPTACRX0 = "Select Distinct Z.*" & vbCrLf _
            & ", Case when VAR_TOLERANCE >= ABS(COST_VAR) THEN 'MATCH' else 'FAIL' END MATCH_FAIL" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR" & vbCrLf _
            & $" from (Select Distinct BOL_NO from {APTACRX1} where PPD_IND = '1') APTACRX1, (" & vbCrLf _
            & "Select BOL_NO" & vbCrLf _
            & ", SUM (NVL(COST_ACT,0)) COST_ACT, SUM (NVL(COST_ACC,0)) COST_ACC, SUM (NVL(TPV_ADJ,0)) TPV_ADJ" & vbCrLf _
            & ", SUM (CASE WHEN NVL(PPD_IND,'0') = '0' THEN 1 ELSE 0 END) LINE_ITEMS" & vbCrLf _
            & ", MAX (VAR_TOLERANCE) VAR_TOLERANCE" & vbCrLf _
            & $" from {APTACRX1} APTACRX1  " & vbCrLf _
            & " group by BOL_NO) X" & vbCrLf _
            & " where APTACRX1.BOL_NO = X.BOL_NO" & vbCrLf _
            & ") Z"
        ASCMAIN1.sql = sqlAPTACRX0 & " where ROWNUM < 1"
        APTACRX0 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTACRX0 & " Add Primary Key (BOL_NO)")

        ASCMAIN1.sql = "Select APTACRX0.* " & vbCrLf _
            & " from " & APTACRX0 & " APTACRX0"
        Create_TDA(dst.Tables.Add, "APTACRX0", "**", 0, False)

        ASCMAIN1.sql = "Select TATTERM1.* " & vbCrLf _
            & " from TATTERM1"
        Create_TDA(dst.Tables.Add, "TATTERM1", "**", 0, False)

        ASCMAIN1.sql = "Select APTACRX1.BOL_NO, APTACRX1.COST_CATGY_CODE, ICTCOST1.ACCT_CODE_TPV" & vbCrLf _
            & ", SUM (APTACRX1.COST_ACC) COST_ACC" & vbCrLf _
            & $" from {APTACRX1} APTACRX1, ICTCOST1" & vbCrLf _
            & " where ICTCOST1.COST_CATGY_CODE = APTACRX1.COST_CATGY_CODE" & vbCrLf _
            & "   and NVL(APTACRX1.PPD_IND,'0') = '0'" & vbCrLf _
            & " group by APTACRX1.BOL_NO, APTACRX1.COST_CATGY_CODE, ICTCOST1.ACCT_CODE_TPV"
        Create_TDA(dst.Tables.Add, "APTACRXA", "**", 0, False,, 3)
        dst.Tables("APTACRXA").Columns.Add("COST_VAR", GetType(System.Decimal))
        dst.Tables("APTACRXA").Columns("COST_VAR").DefaultValue = 0

        ASCMAIN1.sql = "Select APTACRC1.* " & vbCrLf _
            & " from APTACRC1 where ACCRUAL_CODE = 'TRF'" & vbCrLf _
            & " and (CTL_STATUS = '0' or (CTL_STATUS = '1' AND NVL(PPD_IND,'0') = '1' AND NVL(PPD_MATCHED,'0') = '0'))" & vbCrLf _
            & $" and SOURCE_DOC_NO Not in (SELECT BOL_NO FROM {APTACRX0} WHERE MATCH_FAIL = 'MATCH')"
        Create_TDA(dst.Tables.Add, "APTACRC1", "**", 0, False)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        Create_TDA(dst.Tables.Add, "ICTCOST1", "*", 0, False)
        Fill_Records("ICTCOST1")

        ASCMAIN1.sql = $"Select {APTACRCX}.CTL_NO_MATCHED " & vbCrLf &
                $", {APTACRCX}.DEF_TOLERANCE" & vbCrLf &
                $", {APTACRCX}.LNO_COUNT" & vbCrLf &
                $", {APTACRCX}.COST_ACT_TOTAL" & vbCrLf &
                $", {APTACRCX}.COST_ACC_TOTAL" & vbCrLf &
                $", CASE WHEN NVL({APTACRCX}.LNO_COUNT, 0) > 0 THEN NVL({APTACRCX}.COST_ACT_TOTAL, 0) - NVL({APTACRCX}.COST_ACC_TOTAL, 0) ELSE NULL END as CALC_VARIANCE" & vbCrLf &
                $", {APTACRCX}.TPV_ADJ_TOTAL" & vbCrLf &
                $", {APTACRCX}.SOURCE_DOC_NO" & vbCrLf &
                $" from {APTACRCX} " ' & vbCrLf &
        '$" where APTACRC0.ACCRUAL_CODE ='TRF'" & vbCrLf &
        '$" And {APTACRCX}.CTL_NO_MATCHED = APTACRC0.CTL_NO_MATCHED"
        sqlAPTACRCX = ASCMAIN1.sql

        Create_TDA(dst.Tables.Add, "APTACRC0", "**", 0, True)

        With dst.Tables("APTACRC0")
            .Columns("VAR_OK").DefaultValue = "0"
            .Columns.Add("PREV_PPD", GetType(System.Decimal))
            .Columns.Add("PREV_ACC", GetType(System.Decimal))
            .Columns.Add("PREV_VAR", GetType(System.Decimal))
        End With

        ASCMAIN1.sql = "Select * from APTACRC0"
        Create_TDA(dst.Tables.Add, "APTACRCR", "**", 0, False)

        Create_Relation("APTACRC0", "APTACRCR", "CTL_NO_MATCHED", "CTL_NO_MATCHED_NEXT")

        With dst.Tables("APTACRC0")
            .Columns("PREV_PPD").Expression = "SUM(CHILD.COST_ACT_TOTAL)"
            .Columns("PREV_ACC").Expression = "SUM(CHILD.COST_ACC_TOTAL)"
            .Columns("PREV_VAR").Expression = "SUM(CHILD.CALC_VARIANCE)"
        End With



        If perform_fill Then
            Fill_Records_RPT(New Object() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        If parms IsNot Nothing Then
            sqlw = CStr(parms(0))

            'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
            '    Stop
            '    'sqlw &= " and APTACRC1.SOURCE_DOC_NO in ('01669627821','01669627832')"
            '    'sqlw &= " and (APTACRC1.CTL_NO  in ('0000000891','0000000892','0000000788') or APTACRC1.SOURCE_DOC_NO in ('F8636278','F8565808'))"
            '    ' sqlw &= " and (APTACRC1.SOURCE_DOC_NO = 'KQ733348')"
            '    'sqlw &= " and (APTACRC1.SOURCE_DOC_NO like '%KQ224165' or APTACRC1.SOURCE_DOC_NO like '%KQ228786')"
            '    ' 0000004566
            '    sqlw &= " and (APTACRC1.SOURCE_DOC_NO = '540600000101')"
            '    '540600000101
            '    'Else
            '    '    MsgBox("Now Temporarily Removing 3 PPD BOLs ('540500164233','540600000101','SA01231655') with multiple AP Items", MsgBoxStyle.OkOnly, "Temporary Code - this should be removed in April, 2026")
            '    '    sqlw &= " and APTACRC1.SOURCE_DOC_NO NOT IN ('540500164233','540600000101','SA01231655')"
            'End If

            ASCDATA1.ExecuteSQL("Truncate Table " & APTACRX1)
            ASCDATA1.ExecuteSQL("Insert into " & APTACRX1 & " " & Replace(sqlAPTACRX1, " from ", ", NULL MATCH_FAIL from ") & "   and " & sqlw)

            Get_PREV_APTACRC0()

            ASCMAIN1.sql = "Select CASE WHEN PPD_IND = '1' THEN NVL(BOL_NO,SOURCE_DOC_NO) ELSE NVL(BOL_NO,SOURCE_DOC_NO) END SOURCE_DOC_NO" & vbCrLf _
                & ", COUNT (*) RECS" & vbCrLf _
                & ", SUM (CASE WHEN NVL(PPD_IND,'0') = '1' THEN 1 ELSE 0 END) PPDS" & vbCrLf _
                & $" from {APTACRX1} APTACRX1" & vbCrLf _
                & "GROUP BY CASE WHEN PPD_IND = '1' THEN NVL(BOL_NO,SOURCE_DOC_NO) ELSE NVL(BOL_NO,SOURCE_DOC_NO) END" & vbCrLf _
                & "HAVING COUNT (*) > 1"
            ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql}) where PPDS >= 1"

                ASCMAIN1.sql = $"
Begin Declare Cursor C1 is 
 {ASCMAIN1.sql};
 Begin For R1 in C1 Loop
  Update {APTACRX1} Set BOL_NO = R1.SOURCE_DOC_NO
   where (NVL(BOL_NO,SOURCE_DOC_NO) = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO) and BOL_NO is Null;
 End Loop; End;

 Begin Declare Cursor C0 is
  {Replace(ASCMAIN1.sql, "HAVING COUNT (*) > 1", "HAVING COUNT (*) = 1")};
  Begin for R0 in C0 Loop

   Begin Declare Cursor C1 is
    Select * from (Select R0.SOURCE_DOC_NO SOURCE_DOC_NO
    , COUNT (*) RECS
    , SUM (CASE WHEN NVL(PPD_IND,'0') = '1' THEN 1 ELSE 0 END) PPDS
    , MIN (CASE WHEN SOURCE_DOC_NO <> R0.SOURCE_DOC_NO THEN SUBSTR(SOURCE_DOC_NO,1,5) ELSE NULL END) MINSCAC
    , MAX (CASE WHEN SOURCE_DOC_NO <> R0.SOURCE_DOC_NO THEN SUBSTR(SOURCE_DOC_NO,1,5) ELSE NULL END) MAXSCAC
    from {APTACRX1} APTACRX1
    where SOURCE_DOC_NO LIKE '%' || R0.SOURCE_DOC_NO
    HAVING COUNT (*) > 1) where PPDS >= 1 AND MINSCAC = MAXSCAC;
    Begin for R1 in C1 Loop
     Update {APTACRX1} Set BOL_NO = R1.SOURCE_DOC_NO
      where SOURCE_DOC_NO LIKE '%' || R0.SOURCE_DOC_NO and BOL_NO is Null;
    End Loop; End; End;
 End Loop; End; End;

 Delete from {APTACRX1} where BOL_NO is Null;
End;"
            End If
        ASCDATA1.ExecuteSQL()
        ' If ASCMAIN1.Running_in_VS Then Stop
        ' its at this point that you will see the reversals

        ' "Select Distinct Z.*" & vbCrLf _
        ' ", Case when VAR_TOLERANCE >= ABS(COST_VAR) THEN 'MATCH' else 'FAIL' END MATCH_FAIL" & vbCrLf _
        ' " from (" & vbCrLf _
        ' "Select X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR" & vbCrLf _
        ' " from (Select Distinct BOL_NO from ASW74512 where PPD_IND = '1') APTACRX1, (" & vbCrLf _
        ' "Select BOL_NO, COUNT (*) RECS" & vbCrLf _
        ' ", SUM (NVL(COST_ACT,0)) COST_ACT, SUM (NVL(COST_ACC,0)) COST_ACC, SUM (NVL(TPV_ADJ,0)) TPV_ADJ" & vbCrLf _
        ' ", SUM (CASE WHEN NVL(PPD_IND,'0') = '0' THEN 1 ELSE 0 END) LINE_ITEMS" & vbCrLf _
        ' ", MAX (VAR_TOLERANCE) VAR_TOLERANCE" & vbCrLf _
        ' " from ASW74522 APTACRX1" & vbCrLf _
        ' "WHERE SOURCE_DOC_NO IN ('SA01161314','00117570136')" & vbCrLf _
        ' " group by BOL_NO) X" & vbCrLf _
        ' " where APTACRX1.BOL_NO = X.BOL_NO" & vbCrLf _
        ' ") Z" & vbCrLf _
        ' "" & vbCrLf _
        ' BOL_NO               RECS                   COST_ACT               COST_ACC               TPV_ADJ                LINE_ITEMS             VAR_TOLERANCE          COST_VAR               MATCH_FAIL 
        ' -------------------- ---------------------- ---------------------- ---------------------- ---------------------- ---------------------- ---------------------- ---------------------- ---------- 
        ' SA01161314           52                     181071.8               182070.47              0                      50                                            -998.67                FAIL       
        ' REVERSE_0000000216   51                     -182071.8              -182070.47             0                      50                     1.33                   -1.33                  MATCH      
        ' 00117570136          3                      98                     98.01                  0                      1                                             -0.01                  FAIL       
        ' REVERSE_0000000308   2                      -96.8                  -98.01                 0                      1                      1.21                   1.21                   MATCH      
        ' 
        ' 4 Rows Shown of 4 Rows in Result Set"



        '            & $"(SELECT COUNT (DISTINCT ITEM_CODE) * .01  FROM {APTACRX1}" & vbCrLf _

        ASCMAIN1.sql = $"UPDATE {APTACRX1} A SET VAR_TOLERANCE = " & vbCrLf _
            & $"(SELECT COUNT (*) * .01  FROM {APTACRX1}" & vbCrLf _
& "WHERE NVL(BOL_NO,SOURCE_DOC_NO) = A.SOURCE_DOC_NO AND NVL(PPD_IND,'0') = '0') " & vbCrLf _
& " WHERE NVL(VAR_TOLERANCE,0) = 0 AND PPD_IND = '1'"
        ASCDATA1.ExecuteSQL()

        '        Dim AP_PARM_MIN_PPD_MATCH_TOL As Decimal = Val(ROWs("APTPARM1").Item("AP_PARM_MIN_PPD_MATCH_TOL") & "")
        '        ASCMAIN1.sql = $"UPDATE {APTACRX1} A SET VAR_TOLERANCE = {CStr(AP_PARM_MIN_PPD_MATCH_TOL)}" & vbCrLf _
        '& $" WHERE NVL(VAR_TOLERANCE,0) < {CStr(AP_PARM_MIN_PPD_MATCH_TOL)} AND PPD_IND = '1'"
        '        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Insert into {APTACRX0} " & sqlAPTACRX0
        ASCDATA1.ExecuteSQL()

        '        ASCMAIN1.sql = $"UPDATE {APTACRX1} A SET MATCH_FAIL =  (Select MATCH_FAIL from {APTACRX0} where BOL_NO = A.BOL_NO)" & vbCrLf _
        '& $" where PPD_IND = '1'"
        ASCMAIN1.sql = $"UPDATE {APTACRX1} A SET MATCH_FAIL =  (Select MATCH_FAIL from {APTACRX0} where BOL_NO = A.BOL_NO)"
        ASCDATA1.ExecuteSQL()

        '& $" (SELECT CTL_NO_MATCHED from {APTACRX1} APTACRX1 WHERE BOL_NO LIKE 'REVERSE%' AND PPD_IND = '1');" & vbCrLf _
        ASCMAIN1.sql = "" _
& "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
& "SELECT CTL_NO_MATCHED, VAR_TOLERANCE" & vbCrLf _
& " FROM APTACRC0" & vbCrLf _
& " WHERE CTL_NO_MATCHED IN" & vbCrLf _
& $" (Select CTL_NO_MATCHED from APTACRC0 where CTL_NO_MATCHED_NEXT is null and SOURCE_DOC_NO in (Select Distinct SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1'));" & vbCrLf _
& "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
& $"UPDATE {APTACRX1} SET VAR_TOLERANCE = R1.VAR_TOLERANCE, MATCH_FAIL = 'MATCH' WHERE CTL_NO_MATCHED = R1.CTL_NO_MATCHED AND PPD_IND = '1';" & vbCrLf _
& $"UPDATE {APTACRX1} SET MATCH_FAIL = 'MATCH' WHERE CTL_NO_MATCHED = R1.CTL_NO_MATCHED AND PPD_IND = '0';" & vbCrLf _
& "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()


        EnforceConstraints(False)
        Fill_Records("APTACRX1")
        Fill_Records("APTACRX0")
        Fill_Records("APTACRXA")
        Fill_Records("APTACRC1")

        Dim JOURNAL_TYPE = "ICTM"
        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_CTL_DATE As Date = Now.Date
        Dim DETL_POSTING_AMT As Decimal = 0
        Dim ACCT_CODE As String = ""
        Dim rowGLTINTF1 As DataRow

        For Each rowAPTACRX0 As DataRow In dst.Tables("APTACRX0").Select("MATCH_FAIL = 'MATCH'")
            Dim BOL_NO As String = rowAPTACRX0.Item("BOL_NO")

            DETL_POSTING_AMT = -1 * Val(rowAPTACRX0.Item("COST_ACT") & "")
            If DETL_POSTING_AMT <> 0 Then
                rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_TOOLG") & ""
                rowGLTINTF1.Item("DETL_CVX_REF_NO") = BOL_NO
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"
            End If

            DETL_POSTING_AMT = Val(rowAPTACRX0.Item("COST_ACC") & "")
            If DETL_POSTING_AMT <> 0 Then
                rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_ACCR_TOOLG") & ""
                rowGLTINTF1.Item("DETL_CVX_REF_NO") = BOL_NO
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"
            End If

            'If BOL_NO.EndsWith("0101") Then Stop
            Dim COST_VAR As Decimal = Val(rowAPTACRX0.Item("COST_VAR") & "")
            If COST_VAR <> 0 Then

                Dim COST_ACC_TOTAL As Decimal = Val(rowAPTACRX0.Item("COST_ACC") & "")

                ' Spread by BOL/COST_CATGY_CODE

                Dim rowAPTACRXAs() As DataRow = dst.Tables("APTACRXA").Select($"BOL_NO = '{BOL_NO}'", "COST_VAR DESC")
                If rowAPTACRXAs.Length = 1 Then
                    rowAPTACRXAs(0).Item("COST_VAR") = COST_VAR
                Else
                    Dim COST_VAR_TOTAL As Decimal = 0
                    Dim COST_VAR_SNU As Decimal = 0
                    For Each rowAPTACRXA As DataRow In rowAPTACRXAs
                        Dim COST_ACC As Decimal = Val(rowAPTACRXA.Item("COST_ACC") & "")
                        COST_VAR_SNU = System.Math.Round(COST_VAR * COST_ACC / COST_ACC_TOTAL, 2)
                        If COST_VAR_SNU <> 0 Then
                            rowAPTACRXA.Item("COST_VAR") = COST_VAR_SNU
                            COST_VAR_TOTAL += COST_VAR_SNU
                        End If
                    Next
                    If COST_VAR_TOTAL <> COST_VAR Then
                        rowAPTACRXAs(0).Item("COST_VAR") = Val(rowAPTACRXAs(0).Item("COST_VAR") & "") + COST_VAR - COST_VAR_TOTAL
                    End If
                End If

                ' Record J/E by COST_CATGY for this BOL

                Dim COST_VARs As New Dictionary(Of String, Decimal)
                Dim COST_VAR_ADJs As New Dictionary(Of String, Decimal)

                For Each rowAPTACRXA As DataRow In dst.Tables("APTACRXA").Select($"BOL_NO = '{BOL_NO}' and COST_VAR <> 0")
                    DETL_POSTING_AMT = 1 * Val(rowAPTACRXA.Item("COST_VAR") & "")
                    Dim COST_CATGY_CODE As String = rowAPTACRXA.Item("COST_CATGY_CODE")
                    If DETL_POSTING_AMT <> 0 Then
                        rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                        rowGLTINTF1.Item("ACCT_CODE") = rowAPTACRXA.Item("ACCT_CODE_TPV")
                        rowGLTINTF1.Item("DETL_CVX_NO") = COST_CATGY_CODE
                        rowGLTINTF1.Item("DETL_CVX_REF_NO") = BOL_NO
                        rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"

                        COST_VARs.Add(COST_CATGY_CODE, DETL_POSTING_AMT)
                    End If
                Next


                For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select($"BOL_NO = '{BOL_NO}' and TPV_ADJ <> 0")
                    DETL_POSTING_AMT = 1 * Val(rowAPTACRX1.Item("TPV_ADJ") & "")
                    Dim COST_CATGY_CODE As String = rowAPTACRX1.Item("COST_CATGY_CODE")
                    Dim ITEM_CODE As String = rowAPTACRX1.Item("ITEM_CODE")
                    If DETL_POSTING_AMT <> 0 Then
                        rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                        Dim rowICTCOST1 As DataRow = dst.Tables("ICTCOST1").Rows.Find(COST_CATGY_CODE)
                        rowGLTINTF1.Item("ACCT_CODE") = rowICTCOST1.Item("ACCT_CODE_TPV")
                        rowGLTINTF1.Item("DETL_CVX_NO") = COST_CATGY_CODE & ":" & ITEM_CODE
                        rowGLTINTF1.Item("DETL_CVX_REF_NO") = BOL_NO
                        rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"

                        ' COST_VARs.Add(COST_CATGY_CODE, DETL_POSTING_AMT)
                    End If
                Next



                ' Spread by ITEM

                Dim blnIsolate_ADJ As Boolean = True

                For Each SNU As String In COST_VARs.Keys
                    Dim COST_VAR_SNU As Decimal = COST_VARs(SNU)
                    Dim sqlwSNU As String = $"BOL_NO = '{BOL_NO}' and COST_CATGY_CODE = '{SNU}' and ISNULL(PPD_IND,'0') = '0'"

                    Dim COST_VAR_ADJ_TOTAL As Decimal = 0

                    If blnIsolate_ADJ Then
                        Dim rowAPTACRX1_ADJs() As DataRow = dst.Tables("APTACRX1").Select(sqlwSNU & " and ISNULL(TPV_ADJ,0) <> 0")
                        For Each rowAPTACRX1 As DataRow In rowAPTACRX1_ADJs
                            Dim COST_VAR_ITEM As Decimal = Val(rowAPTACRX1.Item("TPV_ADJ") & "")
                            rowAPTACRX1.Item("COST_VAR_ITEM") = COST_VAR_ITEM
                            COST_VAR_ADJ_TOTAL += COST_VAR_ITEM
                        Next
                    End If



                    Dim rowAPTACRX1s() As DataRow = dst.Tables("APTACRX1").Select(sqlwSNU & IIf(blnIsolate_ADJ, " and ISNULL(TPV_ADJ,0) = 0", ""), "COST_ACC DESC")
                    ' COST_VAR_SNU -= COST_VAR_ADJ_TOTAL
                    ' THIS SQL SNIPPET SHOWS WHY YOU DON'T NEED TO REDUCE THE COST_VAR_SNU BY THE TPV_ADJ - TPV_ADJ IS ALREADY SUBTRACTED IN THE CALCULATION OF COST_VAR
                    ' sqlAPTACRX0 = "Select Distinct Z.*" & vbCrLf _
                    '& ", Case when VAR_TOLERANCE >= ABS(COST_VAR) THEN 'MATCH' else 'FAIL' END MATCH_FAIL" & vbCrLf _
                    '& " from (" & vbCrLf _
                    '& "Select X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR" & vbCrLf

                    If rowAPTACRX1s.Length = 0 Then
                        rowAPTACRX1s = dst.Tables("APTACRX1").Select(sqlwSNU, "COST_ACC DESC")
                    End If

                    If rowAPTACRX1s.Length = 1 Then
                        rowAPTACRX1s(0).Item("COST_VAR_ITEM") = COST_VAR_SNU
                    Else
                        Dim COST_ACC_TOTAL_SNU As Decimal = Val(dst.Tables("APTACRX1").Compute("SUM(COST_ACC)", sqlwSNU) & "")
                        Dim COST_VAR_TOTAL As Decimal = 0
                        Dim COST_VAR_ITEM As Decimal = 0
                        For Each rowAPTACRX1 As DataRow In rowAPTACRX1s
                            Dim COST_ACC As Decimal = Val(rowAPTACRX1.Item("COST_ACC") & "")
                            COST_VAR_ITEM = System.Math.Round(COST_VAR_SNU * COST_ACC / COST_ACC_TOTAL_SNU, 2)
                            If COST_VAR_ITEM <> 0 Then
                                rowAPTACRX1.Item("COST_VAR_ITEM") = Val(rowAPTACRX1.Item("COST_VAR_ITEM") & "") + COST_VAR_ITEM
                                COST_VAR_TOTAL += COST_VAR_ITEM
                            End If
                        Next
                        If COST_VAR_TOTAL <> COST_VAR_SNU Then
                            rowAPTACRX1s(0).Item("COST_VAR_ITEM") = Val(rowAPTACRX1s(0).Item("COST_VAR_ITEM") & "") + COST_VAR_SNU - COST_VAR_TOTAL
                        End If
                    End If
                Next
            End If
        Next

        Update_Record_TDA("APTACRX1")






        Dim w As String = "(SOURCE_DOC_NO = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO)"

        ' (SOURCE_DOC_NO = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO)
        ' CTL_NO_MATCHED = R1.SOURCE_DOC_NO
        If w = "CTL_NO_MATCHED = R1.SOURCE_DOC_NO" Then
            w = "CTL_NO_MATCHED = R1.CTL_NO_MATCHED"
        End If
        w = "CTL_NO_MATCHED = R1.CTL_NO_MATCHED" ' NOW THAT WE ARE USING CTL_NO_MATCHED FOR Matched as well as Un-Matched

        '   CTL_NO_MATCHED_gen := TAPCTLN1('APTACRC0.CTL_NO_MATCHED',1);
        ASCMAIN1.sql = $"
Begin Declare CTL_NO_MATCHED_gen VARCHAR2(10); XNO NUMBER(6,0); Cursor C1 is Select Distinct SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1';
 Begin For R1 in C1 Loop
  XNO := NVL(XNO,0) + 1;
  CTL_NO_MATCHED_gen := 'M' || TRIM(TO_CHAR(XNO,'000000000'));
  Update {APTACRX1} Set CTL_NO_MATCHED = CTL_NO_MATCHED_gen where SOURCE_DOC_NO = R1.SOURCE_DOC_NO and PPD_IND = '1';
  Update {APTACRX1} Set CTL_NO_MATCHED = CTL_NO_MATCHED_gen where SOURCE_DOC_NO = R1.SOURCE_DOC_NO and PPD_IND = '0' and CTL_NO_MATCHED is Null;
 End Loop; End;
End;"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Truncate Table " & APTACRCX)
        'ASCDATA1.ExecuteSQL("Insert into " & APTACRCX & " (CTL_NO_MATCHED, SOURCE_DOC_NO) " & "Select CTL_NO_MATCHED, SOURCE_DOC_NO from APTACRC0")
        ASCDATA1.ExecuteSQL($"Insert into {APTACRCX} (CTL_NO_MATCHED, SOURCE_DOC_NO) Select Distinct CTL_NO_MATCHED, SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1'")

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET DEF_TOLERANCE = " & vbCrLf _
            & $"(SELECT COUNT (*) * .01  FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET LNO_COUNT = " & vbCrLf _
            & $"(SELECT COUNT (*) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET COST_ACC_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (COST_ACC) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET COST_ACT_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (COST_ACT) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '1') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET TPV_ADJ_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (TPV_ADJ) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET CALC_VARIANCE = " & vbCrLf _
            & $"Case WHEN NVL(LNO_COUNT, 0) > 0 THEN ROUND(NVL(COST_ACT_TOTAL, 0) - NVL(COST_ACC_TOTAL, 0),2) ELSE NULL END"
        ASCDATA1.ExecuteSQL()


        Fill_Records("APTACRC0")

        Dim rrv0 As DataView = dst.Tables("APTACRC0").DefaultView
        rrv0.RowFilter = "SOURCE_DOC_NO LIKE '*0136'"
        Dim rrt0 As DataTable = rrv0.ToTable

        For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select("PPD_IND = '1' AND BOL_REVERSAL_IND = '1'")
            Dim CTL_NO_MATCHED As String = rowAPTACRX1.Item("CTL_NO_MATCHED") & ""
            Dim SOURCE_DOC_NO As String = rowAPTACRX1.Item("SOURCE_DOC_NO")
            Dim rowAPTACRC0 As DataRow = dst.Tables("APTACRC0").Select($"SOURCE_DOC_NO = '{SOURCE_DOC_NO}'")(0)
            Dim COST_ACT As Decimal = Val(rowAPTACRX1.Item("COST_ACT") & "")
            rowAPTACRC0.Item("COST_ACT_TOTAL") = Round(Val(rowAPTACRC0.Item("COST_ACT_TOTAL") & "") + COST_ACT, 2)
            rowAPTACRC0.Item("CALC_VARIANCE") = Round(Val(rowAPTACRC0.Item("COST_ACT_TOTAL") & "") - Val(rowAPTACRC0.Item("COST_ACC_TOTAL") & ""), 2)
        Next

        Dim rrv00 As DataView = dst.Tables("APTACRC0").DefaultView
        rrv00.RowFilter = "SOURCE_DOC_NO LIKE '*0136'"
        Dim rrt00 As DataTable = rrv00.ToTable

        ' Stop ' COST_ACT_TOTAL IS GETTING RESET HERE

        For Each row As DataRow In dst.Tables("APTACRC0").Select("")
            Dim SOURCE_DOC_NO As String = row.Item("SOURCE_DOC_NO")
            If ASCMAIN1.Running_in_VS AndAlso SOURCE_DOC_NO.EndsWith("0136") Then Stop
            Dim sqlw0 As String = $"PPD_IND = '1' and SOURCE_DOC_NO = '{SOURCE_DOC_NO}' AND ISNULL(BOL_REVERSAL_IND,'0') <> '2'"
            'Dim sqlw0 As String = $"PPD_IND = '1' and SOURCE_DOC_NO = '{SOURCE_DOC_NO}' and BOL_NO NOT LIKE 'REVERSE%'"
            Dim CTL_NO_MATCHED As String = row.Item("CTL_NO_MATCHED")
            Dim VAR_OK As String = dst.Tables("APTACRX1").Compute("MAX(VAR_OK)", sqlw0) & ""
            Dim VAR_TOLERANCE As Decimal = Round(Val(dst.Tables("APTACRX1").Compute("MAX(VAR_TOLERANCE)", sqlw0)), 2)
            Dim COST_ACT_TOTAL As Decimal = Round(Val(dst.Tables("APTACRX1").Compute("SUM(COST_ACT)", sqlw0)), 2)
            Dim NOTES As String = dst.Tables("APTACRX1").Compute("MAX(NOTES)", sqlw0) & ""
            If ASCMAIN1.Running_in_VS AndAlso SOURCE_DOC_NO.EndsWith("0136") Then
                Dim rrvX1 As DataView = dst.Tables("APTACRX1").DefaultView
                rrvX1.RowFilter = "PPD_IND = '1' AND SOURCE_DOC_NO LIKE '*0136'"
                Dim rrtX1 As DataTable = rrvX1.ToTable
                Stop
            End If
            row.Item("PPD_MATCHED_XNO") = XNO
            row.Item("INIT_DATE") = DATETIME_STAMP
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("COST_ACT_TOTAL") = COST_ACT_TOTAL
            row.Item("COST_ACC_TOTAL") = 0
            row.Item("VAR_OK") = VAR_OK
            row.Item("NOTES") = NOTES
            row.Item("VAR_TOLERANCE") = VAR_TOLERANCE
            row.Item("CALC_VARIANCE") = COST_ACT_TOTAL
            row.Item("LNO_COUNT") = 0

            For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select(sqlw0)
                rowAPTACRX1.Item("CTL_NO_MATCHED") = CTL_NO_MATCHED
                row.Item("ACCRUAL_CODE") = rowAPTACRX1.Item("ACCRUAL_CODE")
                Dim OPS_YYYYPP_MATCHED As String = rowAPTACRX1.Item("OPS_YYYYPP_MATCHED") & ""
                If OPS_YYYYPP_MATCHED = "" Then OPS_YYYYPP_MATCHED = ASCMAIN1.CYP
                row.Item("OPS_YYYYPP_MATCHED") = OPS_YYYYPP_MATCHED
                rowAPTACRX1.Item("VAR_TOLERANCE") = VAR_TOLERANCE
                row.Item("TPV_ADJ_TOTAL") = Val(row.Item("TPV_ADJ_TOTAL") & "") + Round(Val(rowAPTACRX1.Item("TPV_ADJ") & ""), 2)

            Next

            sqlw0 = $"PPD_IND = '0' and ISNULL(BOL_NO,SOURCE_DOC_NO) = '{SOURCE_DOC_NO}' AND ISNULL(BOL_REVERSAL_IND,'0') <> '2'"

            For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select(sqlw0)
                row.Item("LNO_COUNT") = Val(row.Item("LNO_COUNT") & "") + 1
                row.Item("COST_ACC_TOTAL") = Val(row.Item("COST_ACC_TOTAL") & "") + Val(rowAPTACRX1.Item("COST_ACC") & "")
                'row.Item("CALC_VARIANCE") = Val(row.Item("CALC_VARIANCE") & "") + Val(rowAPTACRX1.Item("COST_VAR") & "")
                row.Item("TPV_ADJ_TOTAL") = Val(row.Item("TPV_ADJ_TOTAL") & "") + Val(rowAPTACRX1.Item("TPV_ADJ") & "")
            Next

            row.Item("CALC_VARIANCE") = Round(Val(row.Item("COST_ACT_TOTAL") & "") - Val(row.Item("COST_ACC_TOTAL") & ""), 2)
            row.Item("DEF_TOLERANCE") = Round(Val(row.Item("LNO_COUNT") & "") * 0.01, 2)
            ' COULDNT WE USE APTACRCX?
        Next

        For Each row As DataRow In dst.Tables("APTACRC0").Select("")
            Dim SOURCE_DOC_NO As String = row.Item("SOURCE_DOC_NO")
            If ASCMAIN1.Running_in_VS AndAlso SOURCE_DOC_NO.EndsWith("0136") Then Stop
            If ASCMAIN1.Running_in_VS AndAlso SOURCE_DOC_NO.EndsWith("0136") Then
                Dim rrvX1 As DataView = dst.Tables("APTACRX1").DefaultView
                rrvX1.RowFilter = "PPD_IND = '1' AND SOURCE_DOC_NO LIKE '*0136'"
                Dim rrtX1 As DataTable = rrvX1.ToTable


                Dim rrvX0 As DataView = dst.Tables("APTACRC0").DefaultView
                rrvX0.RowFilter = "SOURCE_DOC_NO LIKE '*0136'"
                Dim rrtX0 As DataTable = rrvX0.ToTable
                Stop
            End If
            Dim VAR_OK As String = row.Item("VAR_OK") & ""
            Dim VAR_TOLERANCE As Decimal = Round(Val(row.Item("VAR_TOLERANCE") & ""), 2)
            Dim DEF_TOLERANCE As Decimal = Round(Val(row.Item("DEF_TOLERANCE") & ""), 2)
            Dim CALC_VARIANCE As Decimal = Round(Val(row.Item("CALC_VARIANCE") & ""), 2)
            Dim CALC As Decimal = System.Math.Abs(CALC_VARIANCE)
            Dim LNO_COUNT As Int32 = Val(row.Item("LNO_COUNT") & "")
            If LNO_COUNT >= 1 And (DEF_TOLERANCE >= CALC Or (VAR_OK = "1" And VAR_TOLERANCE >= CALC)) Then
                ' PASS
            Else
                ' FAIL
                row.Delete()
            End If
        Next

        Dim SOURCE_DOC_NOs_reversed As New List(Of String)
        Dim CTL_NO_MATCHEDs_reversed As New List(Of String)
        For Each rowAPTACRCR As DataRow In dst.Tables("APTACRCR").Select("")
            Dim SOURCE_DOC_NO As String = rowAPTACRCR.Item("SOURCE_DOC_NO")
            Dim CTL_NO_MATCHED As String = rowAPTACRCR.Item("CTL_NO_MATCHED")
            SOURCE_DOC_NOs_reversed.Add(SOURCE_DOC_NO)
            CTL_NO_MATCHEDs_reversed.Add(CTL_NO_MATCHED)
            dst.Tables("APTACRC0").Rows.Add(rowAPTACRCR.ItemArray)
        Next

        dst.Tables("APTACRC0").AcceptChanges()

        Dim rrv As DataView = dst.Tables("APTACRX1").DefaultView
        rrv.RowFilter = "PPD_IND = '1' AND SOURCE_DOC_NO LIKE '*0136'"
        Dim rrt As DataTable = rrv.ToTable


        For iteration As Integer = 0 To 1

            For Each row As DataRow In dst.Tables("APTACRC0").Select("", "SOURCE_DOC_NO")

                Dim CTL_NO_MATCHED_prev As String = row.Item("CTL_NO_MATCHED")
                ' do Reversals last
                If CTL_NO_MATCHEDs_reversed.Contains(CTL_NO_MATCHED_prev) Then
                    If iteration = 0 Then Continue For
                Else
                    If iteration = 1 Then Continue For
                End If

                row.SetAdded()

                Dim CTL_NO_MATCHED As String = ASCMAIN1.Next_Control_No("APTACRC0.CTL_NO_MATCHED")
                row.Item("CTL_NO_MATCHED") = CTL_NO_MATCHED

                Dim SOURCE_DOC_NO As String = row.Item("SOURCE_DOC_NO")
                If ASCMAIN1.Running_in_VS AndAlso SOURCE_DOC_NO.EndsWith("0136") Then Stop
                Dim SQLM0 As String = $"PPD_IND = '0' and ISNULL(BOL_NO,SOURCE_DOC_NO) = '{SOURCE_DOC_NO}'"
                Dim SQLM1 As String = $"PPD_IND = '1' and SOURCE_DOC_NO = '{SOURCE_DOC_NO}'"
                If CTL_NO_MATCHEDs_reversed.Contains(CTL_NO_MATCHED_prev) Then
                    SQLM0 = $"PPD_IND = '0' and BOL_NO = 'REVERSE_{CTL_NO_MATCHED_prev}'"
                    SQLM1 = $"PPD_IND = '1' and BOL_NO = 'REVERSE_{CTL_NO_MATCHED_prev}'"

                    ' THS IS BEING HANDLED IN THE UPDATE - BUT THERE ARE FURTHER RESERVATIONS NOTED THERE
                    'Dim isql As String = $"SOURCE_DOC_NO = '{SOURCE_DOC_NO}' and CTL_NO_MATCHED <> '{CTL_NO_MATCHED}'"
                    'Dim rowAPTACRC0_next As DataRow = dst.Tables("APTACRC0").Select(isql)(0)
                    'rowAPTACRC0_next.Item("CTL_NO_MATCHED_PREV") = CTL_NO_MATCHED
                    'row.Item("CTL_NO_MATCHED_NEXT") = rowAPTACRC0_next.Item("CTL_NO")
                End If

                For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select(SQLM1)
                    rowAPTACRX1.Item("CTL_NO_MATCHED") = CTL_NO_MATCHED
                Next
                For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select(SQLM0)
                    rowAPTACRX1.Item("CTL_NO_MATCHED") = CTL_NO_MATCHED
                Next
            Next
        Next



        Update_Record_TDA("APTACRX1")

        ASCMAIN1.sql = "SELECT * FROM (" & vbCrLf _
            & "Select Distinct Z.*" & vbCrLf _
            & ", Case when VAR_TOLERANCE >= ABS(COST_VAR) THEN 'MATCH' else 'FAIL' END MATCH_FAIL" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select X.*, X.COST_ACT - X.COST_ACC - X.TPV_ADJ COST_VAR" & vbCrLf _
            & $" from (Select Distinct BOL_NO from {APTACRX1} where PPD_IND = '1') APTACRX1, (" & vbCrLf _
            & "Select BOL_NO, CTL_NO_MATCHED, SOURCE_DOC_NO, COUNT (*) RECS" & vbCrLf _
            & ", SUM (NVL(COST_ACT,0)) COST_ACT, SUM (NVL(COST_ACC,0)) COST_ACC, SUM (NVL(TPV_ADJ,0)) TPV_ADJ" & vbCrLf _
            & ", SUM (CASE WHEN NVL(PPD_IND,'0') = '0' THEN 1 ELSE 0 END) LINE_ITEMS" & vbCrLf _
            & ", MAX (VAR_TOLERANCE) VAR_TOLERANCE" & vbCrLf _
            & $" from {APTACRX1} APTACRX1" & vbCrLf _
            & "WHERE CTL_NO_MATCHED IS NOT NULL" & vbCrLf _
            & " group by BOL_NO, CTL_NO_MATCHED, SOURCE_DOC_NO) X" & vbCrLf _
            & " where APTACRX1.BOL_NO = X.BOL_NO" & vbCrLf _
            & ") Z" & vbCrLf _
            & $") WHERE MATCH_FAIL = 'FAIL' AND SOURCE_DOC_NO IN (SELECT DISTINCT SOURCE_DOC_NO FROM {APTACRX1} WHERE PPD_IND = '1' AND BOL_REVERSAL_IND = '2' AND MATCH_FAIL = 'MATCH')"

        Dim tblFAIL As DataTable = ASCDATA1.GetDataTable

        For Each rowFAIL As DataRow In tblFAIL.Select("")
            'UPDATE ASW74557 APTACRX1 SET MATCH_FAIL = 'FAIL' WHERE SOURCE_DOC_NO = R1.SOURCE_DOC_NO AND BOL_REVERSAL_IND = '2';
            Dim SOURCE_DOC_NO As String = rowFAIL.Item("SOURCE_DOC_NO")
            For Each rowAPTACRX1 As DataRow In dst.Tables("APTACRX1").Select($"SOURCE_DOC_NO = '{SOURCE_DOC_NO}' AND PPD_IND = '1' AND BOL_REVERSAL_IND = '2' AND MATCH_FAIL = 'MATCH'")
                rowAPTACRX1.Item("MATCH_FAIL") = "FAIL"
                Dim CTL_NO_MATCHED As String = rowAPTACRX1.Item("CTL_NO_MATCHED")
                For Each rowAPTACRX1_ALL As DataRow In dst.Tables("APTACRX1").Select($"CTL_NO_MATCHED = '{CTL_NO_MATCHED}' AND MATCH_FAIL = 'MATCH'")
                    rowAPTACRX1_ALL.Item("MATCH_FAIL") = "FAIL"
                Next
            Next

        Next

        Update_Record_TDA("APTACRX1") ' TO KNOCK OUT THE REVERSAL 2 RECORDS THAT SHOULD NOT GO BECAUSE THE NEW MATCH HAS NOT BEEN OK'D

        'TAC.ICCMAIN1.Prepare_GL_Interface("ICIR", APTACRX1)
        EnforceConstraints(True)
    End Sub

    Function Write_GLTINTF1(
                           JOURNAL_TYPE As String,
                           JOURNAL_NO As String,
                           ByRef JOURNAL_LNO As Integer,
                           DETL_CTL_DATE As Date,
                           DETL_POSTING_AMT As Decimal) As DataRow

        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP0
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ""
        rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
        rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Get_PREV_APTACRC0()

        ASCMAIN1.sql = $"Select Distinct SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1'"
        ASCMAIN1.sql = $"Select * from APTACRC0 where CTL_NO_MATCHED_NEXT is null and SOURCE_DOC_NO in ({ASCMAIN1.sql})"
        Fill_Records("APTACRCR", , , ASCMAIN1.sql)

        Dim r As Integer = 0

        For Each rowAPTACRCR As DataRow In dst.Tables("APTACRCR").Select("")
            Dim CTL_NO_MATCHED_PREV As String = rowAPTACRCR.Item("CTL_NO_MATCHED")
            Dim SOURCE_DOC_NO As String = rowAPTACRCR.Item("SOURCE_DOC_NO")

            rowAPTACRCR.Item("NOTES") = $"Reversing BOL {SOURCE_DOC_NO}, MCTL# {CTL_NO_MATCHED_PREV}"
            rowAPTACRCR.Item("COST_ACC_TOTAL") = -1 * Val(rowAPTACRCR.Item("COST_ACC_TOTAL") & "")
            rowAPTACRCR.Item("COST_ACT_TOTAL") = -1 * Val(rowAPTACRCR.Item("COST_ACT_TOTAL") & "")
            rowAPTACRCR.Item("TPV_ADJ_TOTAL") = -1 * Val(rowAPTACRCR.Item("TPV_ADJ_TOTAL") & "")
            rowAPTACRCR.Item("CALC_VARIANCE") = -1 * Val(rowAPTACRCR.Item("CALC_VARIANCE") & "")

            rowAPTACRCR.Item("CTL_NO_MATCHED_PREV") = CTL_NO_MATCHED_PREV

            ' Create Reversal Records

            ASCMAIN1.sql = $"Select APTACRC1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", NULL QTY_REC, NULL MATCH_FAIL" & vbCrLf _
                & " from APTACRC1, ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf _
                & $" and CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            ASCMAIN1.sql = $"Insert into {APTACRX1} {ASCMAIN1.sql}"

            r = ASCDATA1.ExecuteSQL()
            ' If ASCMAIN1.Running_in_VS Then Stop

            ' & $", COST_ACC = -1 * COST_ACC, COST_ACT = -1 * COST_ACT, TPV_ADJ = -1 * TPV_ADJ" & vbCrLf _
            ASCMAIN1.sql = $"Update {APTACRX1} Set CTL_NO = TAPCTLN1('APTACRC1.CTL_NO',1)" & vbCrLf _
                & $", PPD_MATCHED_XNO = NULL, BOL_NO_MATCHED = NULL, BOL_REVERSAL_IND = '1'" & vbCrLf _
                & $" where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            r = ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Update {APTACRX1} Set CTL_STATUS = '0', OPS_YYYYPP_MATCHED = NULL" & vbCrLf _
                & $" where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'" & vbCrLf _
                & " And PPD_IND = '0'"
            r = ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Update {APTACRX1} Set PPD_MATCHED = NULL" & vbCrLf _
                & $", VAR_TOLERANCE = (Select Count (*) * .01 from {APTACRX1} where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}' And PPD_IND = '0')" & vbCrLf _
                & $" where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'" & vbCrLf _
                & " And PPD_IND = '1' and VAR_TOLERANCE IS NULL"
            r = ASCDATA1.ExecuteSQL()
            ' Stop ' check r

            ASCMAIN1.sql = $"Update {APTACRX1} Set CTL_NO_MATCHED = NULL" & vbCrLf _
                & $" where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            r = ASCDATA1.ExecuteSQL()


            ' Create Accruals for new Match

            ASCMAIN1.sql = $"Select APTACRC1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", NULL QTY_REC, NULL MATCH_FAIL" & vbCrLf _
                & " from APTACRC1, ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf _
                & $" and CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            ASCMAIN1.sql = $"Insert into {APTACRX1} {ASCMAIN1.sql}"

            r = ASCDATA1.ExecuteSQL()
            'If ASCMAIN1.Running_in_VS Then Stop

            '& $", BOL_NO = CASE WHEN PPD_IND = '1' THEN BOL_NO ELSE SOURCE_DOC_NO || '_R' || '{CTL_NO_MATCHED_PREV}' END" & vbCrLf _
            '& $", BOL_NO = SOURCE_DOC_NO || '_R' || '{CTL_NO_MATCHED_PREV}'" & vbCrLf _
            '& $", CTL_NO_MATCHED = NULL" & vbCrLf _

            ' & ", COST_VAR_ITEM = -1 * COST_VAR_ITEM, COST_ACT = -1 * COST_ACT, COST_ACC = -1 * COST_ACC" & vbCrLf _

            '& $", VAR_TOLERANCE = CASE WHEN PPD_IND = '1' THEN (SELECT VAR_TOLERANCE FROM APTACRC0 WHERE CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}') ELSE NULL END" & vbCrLf _

            ASCMAIN1.sql = $"Update {APTACRX1} Set CTL_NO = TAPCTLN1('APTACRC1.CTL_NO',1)" & vbCrLf _
                & $", NOTES = CASE WHEN PPD_IND = '1' THEN 'Re-Matching BOL ' || SOURCE_DOC_NO || ' MCTL# ' || CTL_NO_MATCHED ELSE NOTES END" & vbCrLf _
                & $", BOL_NO = 'REVERSE_' || '{CTL_NO_MATCHED_PREV}', BOL_REVERSAL_IND = '2'" & vbCrLf _
                & $", PPD_MATCHED_XNO = NULL, BOL_NO_MATCHED = NULL" & vbCrLf _
                & ", COST_VAR_ITEM = -1 * COST_VAR_ITEM, COST_ACT = -1 * COST_ACT, COST_ACC = -1 * COST_ACC" & vbCrLf _
                & ", CTL_STATUS = (CASE WHEN NVL(PPD_IND,'0') = '0' THEN '0' ELSE CTL_STATUS END)" & vbCrLf _
                & ", PPD_MATCHED = (CASE WHEN NVL(PPD_IND,'1') = '1' THEN '0' ELSE PPD_MATCHED END)" & vbCrLf _
                & $" where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
            r = ASCDATA1.ExecuteSQL()
            ' If ASCMAIN1.Running_in_VS Then Stop
        Next

    End Sub
End Class