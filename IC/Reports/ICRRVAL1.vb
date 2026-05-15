Public Class ICRRVAL1

    Dim ICTRVAL1 As String
    Dim ICTSTATX As String
    Dim ICTCOSTX As String
    Dim ICTCOSTC As String
    Dim ICTITEMX As String

    Dim ICTCOSTF As String
    Dim ICTVCSTH As String

    Dim ICTRVALI As String

    Dim TABLES As New Dictionary(Of String, String)

    Private JOURNAL_NO As String = String.Empty
    Private Const Unknown As String = "*Unk*"
    Private Const DIFF As String = "ISNULL(PEND_VCOST,0) - ISNULL(CUR_VCOST,0)"
    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables
        MyBase.RWU = "R"

        TABLES = TAC.ICCMAIN1.ReCalculate_Costs(Me, "F", "R")
        ICTSTATX = TABLES("ICTSTATX")
        ICTCOSTX = TABLES("ICTCOSTX")
        ICTITEMX = TABLES("ICTITEMX")

        ICTVCSTH = TABLES("ICTVCSTH")


        'Update_Record_TDA("ICTCOSTX", "1=1")
        ASCDATA1.ExecuteSQL("Truncate Table ICTCOSTX")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTX Select * from " & ICTCOSTX)


        Prepare_Work_File()

        ' Must be Month End and No Filter to Update the Report
        If ASCMAIN1.EOM <> "1" Then
            MyBase.RWU = "N"
        End If

        sql = "Select * from " & ICTRVAL1
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTREVAL", 0))

        sql = "Select * from " & ICTRVALI
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTRVALI", 0))

        sql = "Select * from " & ICTSTATX
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTSTATX", 0))

        sql = "Select * from " & ICTVCSTH
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTVCSTH", 0))
        dst.Tables("ICTVCSTH").Columns.Add("VCOST_DIFF", GetType(System.Decimal), DIFF)

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        ' Make sure all entries have a Valid Cost Category Code to Update to GL
        If MyBase.RWU = "R" Then
            If dst.Tables("ICTREVAL").Select("COST_CATGY_CODE IS NULL OR COST_CATGY_CODE = ''").Length > 0 Then
                MyBase.RWU = "N"
            Else
                For Each rowICTCATG1 As DataRow In ASCDATA1.SelectDistinct("ICTREVAL", "COST_CATGY_CODE").Rows
                    If dst.Tables("ICTCOST1").Select("COST_CATGY_CODE = '" & rowICTCATG1.Item("COST_CATGY_CODE") & "'").Length = 0 Then
                        MyBase.RWU = "N"
                        Exit For
                    End If
                Next
            End If
        End If


        ' If MyBase.RWU = "R" Then
        Prepare_GL_Interface()
        If dst.Tables("ICTRVALG").Select("ACCT_CODE = '" & Unknown & "'").Length > 0 Then
            MyBase.RWU = "N"
        End If
        ' End If

    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("CURR_CODE", ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"))
        CR_params.Add("HDG_LANDG", ROWs("ICTPARM1").Item("IC_PARM_HDG_LANDG"))
        CR_params.Add("HDG_TOOLG", ROWs("ICTPARM1").Item("IC_PARM_HDG_TOOLG"))
        CR_params.Add("HDG_OVRHD", ROWs("ICTPARM1").Item("IC_PARM_HDG_OVRHD"))
        Generate_Report(RPT)

        RPT = "ICRRVAL2"
        Generate_Report(RPT, , "Pending VCost Changes")

        If dst.Tables("GLTINTF1").Rows.Count > 0 Then
            Print_GL()
        End If


        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("ICTRVALI")

        grdASTEXPT1.Text = "Inventory Re-Valuation Variances by Item " & " - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        'Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor Code", 90, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 90, , , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 120, , , Color.LightPink)

        Set_DX_Column(grdASTEXPT1, "QTY_BOM", "Qty BOM", 80, "#,##0", , Color.LightCyan)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOTAL", "Std Cost", 80, "#.0000", , Color.LightCyan)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_FRT_CLASS", "Frt Cls", 60, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TRF_CLASS", "Trf Cls", 60, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY", "MB", 60, , , Color.PaleVioletRed)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_VCOST", "Vnd Direct", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_LANDG", "Frt Direct", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOOLG", "Trf Direct", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_OVRHD", "Ovh Direct", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MATLS", "Vnd Matls", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_LANDGI", "Frt Matls", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOOLGI", "Trf Matls", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_OVRHDI", "Ovh Matls", 80, "#.0000", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 70, , , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70, , , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "SNU", 70, , , Color.LightGray)
        Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70, , , Color.LightGray)

        Set_DX_Column(grdASTEXPT1, "VCOST", "Price $Var", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "MATLS", "Matls $Var", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "LANDG", "Frt $Var", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "TOOLG", "Trf $Var", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "OVRHD", "Ovh $Var", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "TOTAL", "Total $Var", 90, "#,##0", , Color.LightGreen)

        Set_DX_Column(grdASTEXPT1, "LANDGI", "Frt-I $Var", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TOOLGI", "Trf-I $Var", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "OVRHDI", "Ovh_i $Var", 90, "#,##0", , Color.LightBlue)


        Set_DX_Column(grdASTEXPT1, "ITEM_COST_FRT_CLASS_NEW", "Frt Cls New", 60, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TRF_CLASS_NEW", "Trf Cls New", 60, , , Color.PaleVioletRed)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY_NEW", "MB New", 60, , , Color.PaleVioletRed)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_VCOST_NEW", "Vnd Direct New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_LANDG_NEW", "Frt Direct New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOOLG_NEW", "Trf Direct New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_OVRHD_NEW", "Ovh Direct New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MATLS_NEW", "Vnd Matls New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_LANDGI_NEW", "Frt Matls New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_TOOLGI_NEW", "Trf Matls New", 80, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_OVRHDI_NEW", "Ovh Matls New", 80, "#.0000", , Color.Orange)


        For Each C As String In New String() _
            {"VCOST", "MATLS", "LANDG", "TOOLG", "OVRHD", "TOTAL", "LANDGI", "TOOLGI", "OVRHDI"}
            Create_Summary(grdASTEXPT1, C)
        Next

        'grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "BRAND_CODE,COLLECTION_CODE,ITEM_CODE")

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

        ASCMAIN1.sql = "Select ICTVCSTH.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            & $" from {ICTVCSTH} ICTVCSTH, ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = ICTVCSTH.ITEM_CODE"
        Dim tbl2 As DataTable = ASCDATA1.GetDataTable

        tbl2.Columns.Add("VCOST_DIFF", GetType(System.Decimal), DIFF)

        grdASTEXPT2.DataSource = tbl2

        grdASTEXPT2.Text = "Pending VCost Changes"
        tabDataExports.Tabs(1).Visible = True
        tabDataExports.Tabs(1).Text = grdASTEXPT2.Text

        Set_DX_Column(grdASTEXPT2, "")

        ' ASCMAIN1.grdInitializeLayout(grdASTEXPT2)


        Set_DX_Column(grdASTEXPT2, "ITEM_CODE", "Item Code", 120, , "Count", Color.Violet)
        Set_DX_Column(grdASTEXPT2, "ITEM_DESC", "Description", 220, , , Color.Violet)
        Set_DX_Column(grdASTEXPT2, "ITEM_STATUS", "Status", 50, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "COLLECTION_CODE", "Collection", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "PROD_CODE", "Prod Code", 100, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "QTY_OH_PO_IND", "OH/PO Qty Ind", 50, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "ACTIVITY_IND", "Activity Ind", 50, , , Color.CornflowerBlue)
        Set_DX_Column(grdASTEXPT2, "ITEM_RETAIL_PRICE", "Retail Price", 80, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "CUR_VCOST", "VCost", 80, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "PEND_VCOST", "Pending VCost", 80, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "VCOST_DIFF", "Diff", 80, "#,##0.00", , Color.Gold)
        Set_DX_Column(grdASTEXPT2, "CUR_TRF_CLASS", "Trf", 50, , , Color.LightCoral)
        Set_DX_Column(grdASTEXPT2, "PEND_TRF_CLASS", "Pending Trf", 50, , , Color.LightCoral)
        Set_DX_Column(grdASTEXPT2, "CUR_FRT_CLASS", "Frt", 50, , , Color.LightCyan)
        Set_DX_Column(grdASTEXPT2, "PEND_FRT_CLASS", "Pending Frt", 50, , , Color.LightCyan)

        'Set_DX_Column(grdASTEXPT2, "CTL_NO", "Ctl No", 100, , , Color.LightPink)
        'Set_DX_Column(grdASTEXPT2, "CTL_DATE", "Date", 100, "MM/dd/yy", , Color.LightPink)
        'Set_DX_Column(grdASTEXPT2, "CTL_NOTE", "Note", 200, , , Color.LightPink)

        'Set_DX_Column(grdASTEXPT2, "COST_ACC", "Accrued", 120, "#,##0.00", "Sum", Color.LightGreen)
        'Set_DX_Column(grdASTEXPT2, "TPV_ADJ", "TPV Adj", 120, "#,##0.00", "Sum", Color.LightGreen)
        'Set_DX_Column(grdASTEXPT2, "COST_ACT", "Pre-Paid", 120, "#,##0.00", "Sum", Color.LightGreen)
        'Set_DX_Column(grdASTEXPT2, "COST_VAR", "Variance", 120, "#,##0.00", "Sum", Color.LightGreen)

        'Set_DX_Column(grdASTEXPT2, "OPS_YYYYPP", "YP", 75, , , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT2, "VOUCHER_NO", "Voucher No", 100, , , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT2, "RECEIPT_NO", "Receipt No", 70, , , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT2, "RECEIPT_LNO", "Ln", 30, "##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT2, "PO_ORDER_NO", "PO No", 70, , , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT2, "PO_ORDER_LNO", "Ln", 30, "##0", , Color.LightBlue)

        'Set_DX_Column(grdASTEXPT2, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        'Set_DX_Column(grdASTEXPT2, "ITEM_DESC", "Description", 130, , , Color.Gold)
        'Set_DX_Column(grdASTEXPT2, "QTY_REC", "Qty Rec", 90, "#,##0", "Sum", Color.LightBlue)

        'Sort_grdColumns(grdASTEXPT2, "SOURCE_DOC_NO")

        grdASTEXPT2.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT2, "ITEM_CODE")


    End Sub

    Sub Prepare_Work_File()

        If ICTRVAL1 = "" Then
            ASCMAIN1.sql = Get_SQL(False)
            ICTRVAL1 = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & ICTRVAL1 & " Add Primary Key (BRAND_CODE, COLLECTION_CODE, COST_CATGY_CODE, PROD_CODE, ITEM_CATGY_EXP)")
            ASCMAIN1.AnalyzeTable(ICTRVAL1)

            ASCMAIN1.sql = Get_SQL(True)
            ICTRVALI = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL("Alter Table " & ICTRVALI & " Add Primary Key (ITEM_CODE)")
            ASCMAIN1.AnalyzeTable(ICTRVALI)

        Else
            ASCMAIN1.sql = Get_SQL(False)
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTRVAL1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTRVAL1 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = Get_SQL(True)
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTRVALI)
            ASCDATA1.ExecuteSQL("Insert into " & ICTRVALI & " " & ASCMAIN1.sql)
        End If

    End Sub

    Function Get_SQL(detailed As Boolean) As String

        Dim sqlICTRVALI As String = "" _
            & vbCrLf & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEMX.QTY_BOM, ICTCOSTC.ITEM_COST_TOTAL" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_FRT_CLASS, ICTCOSTC.ITEM_COST_TRF_CLASS, ICTCOSTC.ITEM_COST_MAKE_BUY" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_VCOST, ICTCOSTC.ITEM_COST_LANDG, ICTCOSTC.ITEM_COST_TOOLG, ICTCOSTC.ITEM_COST_OVRHD" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_MATLS, ICTCOSTC.ITEM_COST_LANDGI, ICTCOSTC.ITEM_COST_TOOLGI, ICTCOSTC.ITEM_COST_OVRHDI" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_FRT_CLASS ITEM_COST_FRT_CLASS_NEW, ICTCOSTX.ITEM_COST_TRF_CLASS ITEM_COST_TRF_CLASS_NEW, ICTCOSTX.ITEM_COST_MAKE_BUY ITEM_COST_MAKE_BUY_NEW" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_VCOST ITEM_COST_VCOST_NEW, ICTCOSTX.ITEM_COST_LANDG ITEM_COST_LANDG_NEW, ICTCOSTX.ITEM_COST_TOOLG ITEM_COST_TOOLG_NEW, ICTCOSTX.ITEM_COST_OVRHD ITEM_COST_OVRHD_NEW" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_MATLS ITEM_COST_MATLS_NEW, ICTCOSTX.ITEM_COST_LANDGI ITEM_COST_LANDGI_NEW, ICTCOSTX.ITEM_COST_TOOLGI ITEM_COST_TOOLGI_NEW, ICTCOSTX.ITEM_COST_OVRHDI ITEM_COST_OVRHDI_NEW"

        Dim sqlICTRVALI_group_by As String = "" _
            & vbCrLf & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEMX.QTY_BOM, ICTCOSTC.ITEM_COST_TOTAL" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_FRT_CLASS, ICTCOSTC.ITEM_COST_TRF_CLASS, ICTCOSTC.ITEM_COST_MAKE_BUY" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_VCOST, ICTCOSTC.ITEM_COST_LANDG, ICTCOSTC.ITEM_COST_TOOLG, ICTCOSTC.ITEM_COST_OVRHD" _
            & vbCrLf & ", ICTCOSTC.ITEM_COST_MATLS, ICTCOSTC.ITEM_COST_LANDGI, ICTCOSTC.ITEM_COST_TOOLGI, ICTCOSTC.ITEM_COST_OVRHDI" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_FRT_CLASS, ICTCOSTX.ITEM_COST_TRF_CLASS, ICTCOSTX.ITEM_COST_MAKE_BUY" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_VCOST, ICTCOSTX.ITEM_COST_LANDG, ICTCOSTX.ITEM_COST_TOOLG, ICTCOSTX.ITEM_COST_OVRHD" _
            & vbCrLf & ", ICTCOSTX.ITEM_COST_MATLS, ICTCOSTX.ITEM_COST_LANDGI, ICTCOSTX.ITEM_COST_TOOLGI, ICTCOSTX.ITEM_COST_OVRHDI"


        If Not detailed Then sqlICTRVALI = ""
        If Not detailed Then sqlICTRVALI_group_by = ""

        Dim sqlICTRVAL1 As String = "Select ICTITEMX.BRAND_CODE, ICTITEMX.COLLECTION_CODE, ICTITEMX.COST_CATGY_CODE, ICTITEMX.PROD_CODE, NVL(ICTCOST1.EXP_AT_PURCHASE,'0') ITEM_CATGY_EXP" _
            & sqlICTRVALI & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_TOTAL,0) - NVL(ICTCOSTC.ITEM_COST_TOTAL,0)) * NVL(ICTITEMX.QTY_BOM,0)) TOTAL" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_VCOST,0) - NVL(ICTCOSTC.ITEM_COST_VCOST,0)) * NVL(ICTITEMX.QTY_BOM,0)) VCOST" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_MATLS,0) - NVL(ICTCOSTC.ITEM_COST_MATLS,0)) * NVL(ICTITEMX.QTY_BOM,0)) MATLS" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_LANDG,0) - NVL(ICTCOSTC.ITEM_COST_LANDG,0)) * NVL(ICTITEMX.QTY_BOM,0)) LANDG" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_LANDGI,0) - NVL(ICTCOSTC.ITEM_COST_LANDGI,0)) * NVL(ICTITEMX.QTY_BOM,0)) LANDGI" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_TOOLG,0) - NVL(ICTCOSTC.ITEM_COST_TOOLG,0)) * NVL(ICTITEMX.QTY_BOM,0)) TOOLG" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_TOOLGI,0) - NVL(ICTCOSTC.ITEM_COST_TOOLGI,0)) * NVL(ICTITEMX.QTY_BOM,0)) TOOLGI" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_OVRHD,0) - NVL(ICTCOSTC.ITEM_COST_OVRHD,0)) * NVL(ICTITEMX.QTY_BOM,0)) OVRHD" & vbCrLf _
            & ", Sum ((NVL(ICTCOSTX.ITEM_COST_OVRHDI,0) - NVL(ICTCOSTC.ITEM_COST_OVRHDI,0)) * NVL(ICTITEMX.QTY_BOM,0)) OVRHDI" & vbCrLf _
            & " from " & ICTCOSTX & " ICTCOSTX, ICTCOSTC, " & ICTITEMX & " ICTITEMX, ICTCOST1, ICTITEM1" & vbCrLf _
            & " where ICTCOSTC.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf _
            & "   and ICTITEMX.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf _
            & "   and ICTCOST1.COST_CATGY_CODE = ICTITEMX.COST_CATGY_CODE" & vbCrLf _
            & "   and NVL(ICTITEMX.QTY_BOM,0) <> 0" & vbCrLf _
            & " group by ICTITEMX.BRAND_CODE, ICTITEMX.COLLECTION_CODE, ICTITEMX.COST_CATGY_CODE, ICTITEMX.PROD_CODE, NVL(ICTCOST1.EXP_AT_PURCHASE,'0')" _
            & sqlICTRVALI_group_by

        Return sqlICTRVAL1
    End Function

    Overrides Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Updating Cost Files")

        ASCMAIN1.sql = "" _
          & "Begin" & vbCrLf _
          & " Declare Cursor C1 IS " & vbCrLf _
          & "   Select ICTCOSTX.*, ICTITEM1.ITEM_YYYYPP_CUR_COST, ICTITEM1.ITEM_COST_STD" & vbCrLf _
          & "   , NVL(ICTITEMX.QTY_BOM,0) * (NVL(ICTCOSTX.ITEM_COST_TOTAL,0) - NVL(ICTCOSTC.ITEM_COST_TOTAL,0)) TOTAL" & vbCrLf _
          & "    from ICTCOSTX,ICTCOSTC," & ICTITEMX & " ICTITEMX,ICTITEM1" & vbCrLf _
          & "     where ICTCOSTC.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf _
          & "       and ICTITEMX.ITEM_CODE = ICTCOSTX.ITEM_CODE" & vbCrLf _
          & "       and ICTITEM1.ITEM_CODE = ICTCOSTX.ITEM_CODE;" & vbCrLf _
          & " Begin " & vbCrLf _
          & "   For R1 IN C1 Loop" & vbCrLf _
          & "    If R1.ITEM_COST_MAKE_BUY = 'M' Then" & vbCrLf _
          & "     Update BMTMAIN1 " & vbCrLf _
          & "       Set BM_ISSUE_STD = R1.BM_ISSUE_NO" & vbCrLf _
          & "       where BM_PROD_ITEM = R1.ITEM_CODE;" & vbCrLf _
          & "    End If;" & vbCrLf _
          & "    If NVL(R1.TOTAL,0) <> 0 Then" & vbCrLf _
          & "     Update ICTIVAR1 Set RV_EXP = NVL(RV_EXP,0) - NVL(R1.TOTAL,0)" & vbCrLf _
          & "      where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
          & "     If SQL%NOTFOUND Then " & vbCrLf _
          & "      Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, RV_EXP) " & vbCrLf _
          & "       Values (R1.ITEM_CODE, '" & ASCMAIN1.CYP & "', -1 * R1.TOTAL);" & vbCrLf _
          & "     End If;" & vbCrLf _
          & "    End If;" & vbCrLf _
          & "    Update ICTITEM1 Set ITEM_COST_STD = R1.ITEM_COST_TOTAL" & vbCrLf _
          & "     , ITEM_COST_CURR_CODE = R1.ITEM_COST_CURR_CODE" & vbCrLf _
          & "     , ITEM_COST_MAKE_BUY = R1.ITEM_COST_MAKE_BUY" & vbCrLf _
          & "     , ITEM_PLAN_MAKE_BUY = R1.ITEM_COST_MAKE_BUY" & vbCrLf _
          & "     , ITEM_COST_FRT_CLASS = R1.ITEM_COST_FRT_CLASS" & vbCrLf _
          & "     , ITEM_COST_TRF_CLASS = R1.ITEM_COST_TRF_CLASS" & vbCrLf _
          & "     , ITEM_COST_WASTE_PCT = R1.ITEM_COST_WASTE_PCT" & vbCrLf _
          & "     , ITEM_YYYYPP_PRV_COST = R1.ITEM_YYYYPP_CUR_COST" & vbCrLf _
          & "     , ITEM_YYYYPP_CUR_COST = '" & ASCMAIN1.CYP & "'" & vbCrLf _
          & "     where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
          & "   End Loop; " & vbCrLf _
          & " End; " & vbCrLf _
          & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "" _
          & "Begin" & vbCrLf _
          & " Declare Cursor C1 IS " & vbCrLf _
          & "   Select ICTRVALI.*" & vbCrLf _
          & $"    from {ICTRVALI} ICTRVALI, {ICTITEMX} ICTITEMX" & vbCrLf _
          & "     where ICTITEMX.ITEM_CODE = ICTRVALI.ITEM_CODE" & vbCrLf _
          & "       and (NVL(ICTRVALI.VCOST,0) <> 0 or NVL(ICTRVALI.MATLS,0) <> 0" & vbCrLf _
          & "         or NVL(ICTRVALI.LANDG,0) <> 0 or NVL(ICTRVALI.LANDGI,0) <> 0" & vbCrLf _
          & "         or NVL(ICTRVALI.TOOLG,0) <> 0 or NVL(ICTRVALI.TOOLGI,0) <> 0" & vbCrLf _
          & "         or NVL(ICTRVALI.OVRHD,0) <> 0 or NVL(ICTRVALI.OVRHDI,0) <> 0);" & vbCrLf _
          & " Begin " & vbCrLf _
          & "  For R1 IN C1 Loop" & vbCrLf _
          & "   Update ICTIVAR1 Set" & vbCrLf _
          & "     RV_EXP_VCOST = NVL(RV_EXP_VCOST,0) - NVL(R1.VCOST,0)" & vbCrLf _
          & ",    RV_EXP_LANDG = NVL(RV_EXP_LANDG,0) - NVL(R1.LANDG,0)" & vbCrLf _
          & ",    RV_EXP_TOOLG = NVL(RV_EXP_TOOLG,0) - NVL(R1.TOOLG,0)" & vbCrLf _
          & ",    RV_EXP_OVRHD = NVL(RV_EXP_OVRHD,0) - NVL(R1.OVRHD,0)" & vbCrLf _
          & ",    RV_EXP_MATLS = NVL(RV_EXP_MATLS,0) - NVL(R1.MATLS,0)" & vbCrLf _
          & ",    RV_EXP_LANDGI = NVL(RV_EXP_LANDGI,0) - NVL(R1.LANDGI,0)" & vbCrLf _
          & ",    RV_EXP_TOOLGI = NVL(RV_EXP_TOOLGI,0) - NVL(R1.TOOLGI,0)" & vbCrLf _
          & ",    RV_EXP_OVRHDI = NVL(RV_EXP_OVRHDI,0) - NVL(R1.OVRHDI,0)" & vbCrLf _
          & "    where ITEM_CODE = R1.ITEM_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
          & "   If SQL%NOTFOUND Then " & vbCrLf _
          & "    Insert into ICTIVAR1 (ITEM_CODE, OPS_YYYYPP, RV_EXP_VCOST, RV_EXP_LANDG, RV_EXP_TOOLG, RV_EXP_OVRHD, RV_EXP_MATLS, RV_EXP_LANDGI, RV_EXP_TOOLGI, RV_EXP_OVRHDI) " & vbCrLf _
          & "     Values (R1.ITEM_CODE, '" & ASCMAIN1.CYP & "', -1 * R1.VCOST, -1 * R1.LANDG, -1 * R1.TOOLG, -1 * R1.OVRHD, -1 * R1.MATLS, -1 * R1.LANDGI, -1 * R1.TOOLGI, -1 * R1.OVRHDI);" & vbCrLf _
          & "   End If;" & vbCrLf _
          & "  End Loop; " & vbCrLf _
          & " End; " & vbCrLf _
          & "End;"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Update ICTITEM1 Set ITEM_PLAN_WASTE_PCT = (Select ITEM_WASTE_PCT from ICTTYPE1 where ITEM_TYPE_CODE = ICTITEM1.ITEM_TYPE_CODE)")
        ASCDATA1.ExecuteSQL("Update ICTITEM1 Set ITEM_COST_WASTE_PCT = (Select ITEM_WASTE_PCT from ICTTYPE1 where ITEM_TYPE_CODE = ICTITEM1.ITEM_TYPE_CODE)")
    
        ASCDATA1.ExecuteSQL("Delete from ICTCOSTC where ITEM_CODE in (Select ITEM_CODE from ICTCOSTX)")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTC Select * from ICTCOSTX")
        ASCDATA1.ExecuteSQL("Delete from ICTCOSTH where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and ITEM_CODE in (Select ITEM_CODE from ICTCOSTX)")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTH Select '" & ASCMAIN1.CYP & "', ICTCOSTX.* from ICTCOSTX")
        ASCDATA1.ExecuteSQL("Delete from ICTCOSTA where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and ITEM_CODE in (Select ITEM_CODE from ICTCOSTX)")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTA Select '" & ASCMAIN1.CYP & "', ICTCOSTX.* from ICTCOSTX")
        ASCDATA1.ExecuteSQL("Update ICTFRTC1 Set FRT_CLASS_PCT_CUR = FRT_CLASS_PCT_FUT")
        ASCDATA1.ExecuteSQL("Update ICTTRFC1 Set TRF_CLASS_PCT_CUR = TRF_CLASS_PCT_FUT")
        ASCDATA1.ExecuteSQL($"INSERT INTO smz_ICTCOSTF_jic select ictcostf.*, '{XNO}' XNO, sysdate XNO_DATE from ictcostf")
        ASCDATA1.ExecuteSQL("Delete from ICTCOSTF")
        ASCDATA1.ExecuteSQL($"Insert into ICTCOSTF Select * from ICTCOSTX")

        'ASCDATA1.ExecuteSQL("Update ICTCATG1 Set OVERHEAD_PCT_CUR = OVERHEAD_PCT_FUT")
        'ASCDATA1.ExecuteSQL("Update TATCURR1 Set CURR_EXCH_CUR = CURR_EXCH_FUT")
        'ASCDATA1.ExecuteSQL("Delete from ICTVCSTH where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'")
        'ASCDATA1.ExecuteSQL($"INSERT INTO ICTVCSTH(OPS_YYYYPP, ITEM_CODE, VEND_CODE, COLLECTION_CODE
        ', ITEM_STATUS, PROD_CODE, ITEM_RETAIL_PRICE, VCOST, CALC_VCOST, QTY_OH_PO_IND, ACTIVITY_IND
        ', COST_LIST_PO_VCOST) Select OPS_YYYYPP, ITEM_CODE, VEND_CODE, COLLECTION_CODE
        ', ITEM_STATUS, PROD_CODE, ITEM_RETAIL_PRICE, VCOST, CALC_VCOST, QTY_OH_PO_IND, ACTIVITY_IND
        ', COST_LIST_PO_VCOST from {ICTVCSTH}")

        ' SHOULD WE BE SETTING PLAN MB & PLAN WASTE BELOW ?

        ' do not update plan waste % with std cost waste % - keep independently maintained
        'dynICTITEM1.Fields("ITEM_PLAN_WASTE_PCT").Value = dynICWCOSTX.Fields("ITEM_COST_WASTE_PCT").Value

        ' Catch Items whose cost was created this period, and perhaps re-initialized, after a few transactions were entered

        TAC.ICCMAIN1.Update_Movement_Costs()
        GL_Update()


        ASCMAIN1.sql = "SELECT * FROM (" & vbCrLf _
            & "SELECT I.ITEM_CODE, I.ITEM_COST_STD, C.ITEM_COST_TOTAL, I.ITEM_COST_STATUS" & vbCrLf _
            & "FROM ICTITEM1 I, ICTCOSTC C" & vbCrLf _
            & "WHERE C.ITEM_CODE (+) = I.ITEM_CODE" & vbCrLf _
            & ") WHERE (ITEM_COST_STATUS = 'P' AND ITEM_COST_TOTAL IS NOT NULL)" & vbCrLf _
            & "OR (ITEM_COST_STD <> ITEM_COST_TOTAL)"
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        If TBL.Rows.Count > 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(TBL, Me, "Cost Database Issues ICTITEM1 vs ICTCOSTC - please contact ABS")
            End Using
        End If
        If ASCMAIN1.Running_in_VS Then
            'Stop
            'ASCMAIN1.sql = "select * from ictcostf"
            'Dim tbl2 As DataTable = ASCDATA1.GetDataTable()
            'ASCMAIN1.sql = "select count(*) from ictcostf"
            'Dim tbl3 As DataTable = ASCDATA1.GetDataTable()
            'ASCMAIN1.sql = "select * from ictcostc where item_code='JL004P80'"
            'Dim tbl4 As DataTable = ASCDATA1.GetDataTable()
            'Stop
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub Prepare_GL_Interface()

        ' Prepare GL Interface File

        dst.Tables("GLTINTF1").Clear()

        JOURNAL_NO = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0
        Dim JOURNAL_TYPE As String = "ICRV"

        Dim DETL_POSTING_AMT As Double = 0
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))

        ' Note - the placeholders for {0} and {1} are replaced down below
        Dim sql As String = "Select ICTRVAL1.COST_CATGY_CODE " & vbCrLf _
            & ", ICTCOST1.{0} ACCT_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ", NVL(ICTCOLL1.SEG4_CODE,ICTRVAL1.COLLECTION_CODE) SEG4_CODE" & vbCrLf _
            & ", 'ADJ' DIST_TYPE, ICTRVAL1.DIST_AMT" & vbCrLf _
            & " from ICTCOST1, ICTCOLL1, (Select COST_CATGY_CODE, COLLECTION_CODE, Sum ({1}) DIST_AMT from " & ICTRVAL1 & " group by COST_CATGY_CODE, COLLECTION_CODE) ICTRVAL1" & vbCrLf _
            & " where ICTCOST1.COST_CATGY_CODE = ICTRVAL1.COST_CATGY_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTRVAL1.COLLECTION_CODE" & vbCrLf _
            & "   and NVL(ICTRVAL1.DIST_AMT,0) <> 0"

        Dim sqlSEG2 As String = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE"
        If ASCMAIN1.CLIENT = "AHA" Then
            sqlSEG2 = "'GNA' SEG2_CODE"
        End If

        ASCMAIN1.sql = "" _
            & Replace(Replace(Replace(Replace(String.Format(sql, "ACCT_CODE_ONH", "TOTAL"),
                       "NVL(ICTCOLL1.SEG4_CODE,ICTRVAL1.COLLECTION_CODE)",
                       "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'"),
                   "COST_CATGY_CODE, COLLECTION_CODE",
                   "COST_CATGY_CODE"),
                "ICTCOST1, ICTCOLL1,", "ICTCOST1,"),
                "   and ICTCOLL1.COLLECTION_CODE = ICTRVAL1.COLLECTION_CODE", "") _
            & " union " & vbCrLf _
            & Replace(String.Format(sql, "ACCT_CODE_RVAL_B", "-1  * (VCOST)"), "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE", sqlSEG2) _
            & " union " & vbCrLf _
            & Replace(String.Format(sql, "ACCT_CODE_RVAL_M", "-1  * (MATLS)"), "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE", sqlSEG2) _
            & " union " & vbCrLf _
            & Replace(String.Format(sql, "ACCT_CODE_RVAL_F", "-1  * (LANDG + LANDGI)"), "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE", sqlSEG2) _
            & " union " & vbCrLf _
            & Replace(String.Format(sql, "ACCT_CODE_RVAL_T", "-1  * (TOOLG + TOOLGI)"), "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE", sqlSEG2)

        ' NOTE: really should have supported OVRHD as it's own category
        '     - and may need to do this if we ever need this std cost bucket
        '     - at that time we will need to support ICTCOST1.ACCT_CODE_RVAL_O
        '& " union " & vbCrLf _
        '& Replace(String.Format(sql, "ACCT_CODE_RVAL_O", "-1  * (OVRHD + OVRHDI)"), "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE", sqlSEG2)

        ' NOTE: MATLS is really VCOSTI

        Dim ICTRVALG As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("UPDATE " & ICTRVALG & " SET ACCT_CODE = '" & Unknown & "' WHERE ACCT_CODE IS NULL")

        If ASCMAIN1.CLIENT = "AHA" Then
            ASCDATA1.ExecuteSQL("UPDATE " & ICTRVALG & " SET SEG2_CODE = '000' where ACCT_CODE = '6162'")
        End If

        sql = "Select * from " & ICTRVALG
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTRVALG", 0))

        For Each row As DataRow In dst.Tables("ICTRVALG").Rows
            DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")
            Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = ASCMAIN1.CYP ' ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = row("ACCT_CODE")
            rowGLTINTF1("SEG2_CODE") = row("SEG2_CODE")
            rowGLTINTF1("SEG3_CODE") = row("SEG3_CODE")
            rowGLTINTF1("SEG4_CODE") = row("SEG4_CODE")
            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
            rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
            rowGLTINTF1("DETL_EXE_NO") = XNO
            rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
            rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_NO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
            rowGLTINTF1("DETL_CVX_REF_NO") = DBNull.Value
            rowGLTINTF1("DETL_DESC") = DBNull.Value
            rowGLTINTF1("DETL_CVX_TYPE") = DBNull.Value
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Next
    End Sub
End Class