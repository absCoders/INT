Public Class ICRIVAL1
    Dim ICTIVAL1 As String
    Dim RYP_LM As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("ICTPARM1")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Prepare_Work_File()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.sql = "Select ICTITEM1.* " _
            & ", DECODE(ICTCOST1.EXP_AT_PURCHASE,'1'," & IIf(Absx1.chkFor("VALUE_EXP_AS_PURCH_AS_ZERO").Checked, 0, "ICTCOSTA_LM.ITEM_COST_TOTAL") & ",ICTCOSTA_LM.ITEM_COST_TOTAL) STD_COST_LM" _
            & ", DECODE(ICTCOST1.EXP_AT_PURCHASE,'1'," & IIf(Absx1.chkFor("VALUE_EXP_AS_PURCH_AS_ZERO").Checked, 0, "ICTCOSTA_TM.ITEM_COST_TOTAL") & ",ICTCOSTA_TM.ITEM_COST_TOTAL) STD_COST_TM" _
            & ", ICTCOSTA_TM.ITEM_COST_VCOST" & vbCrLf _
            & ", ICTCOSTA_TM.ITEM_COST_MATLS" & vbCrLf _
            & ", ICTCOSTA_TM.ITEM_COST_LANDGI" & vbCrLf _
            & ", ICTCOSTA_TM.ITEM_COST_TOOLGI" & vbCrLf _
            & ", ICTCOSTA_TM.ITEM_COST_LANDG" & vbCrLf _
            & ", ICTCOSTA_TM.ITEM_COST_TOOLG" & vbCrLf _
            & " from ICTITEM1, ICTCOSTA ICTCOSTA_LM, ICTCOSTA ICTCOSTA_TM, ICTCOST1" _
            & " where ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & ICTIVAL1 & ")" _
            & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE" _
            & "   and ICTCOSTA_LM.OPS_YYYYPP (+) = '" & RYP_LM & "'" _
            & "   and ICTCOSTA_LM.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" _
            & "   and ICTCOSTA_TM.OPS_YYYYPP (+) = '" & RYP & "'" _
            & "   and ICTCOSTA_TM.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTITEM1", 1))

        Call MyBase.Get_SQL("*", ICTIVAL1)

        Dim SOURCE_TABLE_NAME As String = "ICTIVAL1"
        Dim x As String = ASTSRPT1_sum_columns
        Dim y As String = ASTSRPT1_sql_sum
        Dim sql_Data As String = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", " & SOURCE_TABLE_NAME & ".ITEM_CODE" _
        & ASTSRPT1_sum_columns _
        & " from " & ICTIVAL1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols _
        & ", " & SOURCE_TABLE_NAME & ".ITEM_CODE"

        'sql = "Select " & sql_SELECT_cols & vbCrLf _
        '& ", " & SOURCE_TABLE_NAME & ".ITEM_CODE" _
        '& " from " & ICTIVAL1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        '& ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        '& " group by " & sql_GROUP_BY_cols _
        '& ", " & SOURCE_TABLE_NAME & ".ITEM_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        'ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        '& "(" & G1thru9 & COLUMN_NAMEs_appended & DATA_COLUMN_NAMEs & ")" & vbCrLf _
        '& "(" & sql & ")"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = RYPLEGEND
        CR_params.Add("OPTDS", IIf(Absx1.chkFor("SUPPRESS_ITEM_DETAIL").Checked, "S", "D"))
        Generate_Report(RPT, , SUBT)

        Prepare_Data_Extracts()
    End Sub


    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        With dst.Tables("ASTSRPT1").Columns
            .Add("ITEM_DESC")
            .Add("COLLECTION_CODE")
            .Add("ITEM_COST_MAKE_BUY")
            .Add("COST_CATGY_CODE")
            .Add("PROD_CODE")
            .Add("VEND_CODE")
            .Add("STD_COST_LM", GetType(System.Decimal))
            .Add("STD_COST_TM", GetType(System.Decimal))
            .Add("ITEM_COST_VCOST", GetType(System.Decimal))
            .Add("ITEM_COST_MATLS", GetType(System.Decimal))
            .Add("ITEM_COST_LANDGI", GetType(System.Decimal))
            .Add("ITEM_COST_TOOLGI", GetType(System.Decimal))
            .Add("ITEM_COST_LANDG", GetType(System.Decimal))
            .Add("ITEM_COST_TOOLG", GetType(System.Decimal))
        End With

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            row.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
            row.Item("ITEM_COST_MAKE_BUY") = rowICTITEM1.Item("ITEM_COST_MAKE_BUY")
            row.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
            row.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
            row.Item("VEND_CODE") = rowICTITEM1.Item("VEND_CODE")
            row.Item("STD_COST_LM") = rowICTITEM1.Item("STD_COST_LM")
            row.Item("STD_COST_TM") = rowICTITEM1.Item("STD_COST_TM")
            row.Item("ITEM_COST_VCOST") = rowICTITEM1.Item("ITEM_COST_VCOST")
            row.Item("ITEM_COST_MATLS") = rowICTITEM1.Item("ITEM_COST_MATLS")
            row.Item("ITEM_COST_LANDGI") = rowICTITEM1.Item("ITEM_COST_LANDGI")
            row.Item("ITEM_COST_TOOLGI") = rowICTITEM1.Item("ITEM_COST_TOOLGI")
            row.Item("ITEM_COST_LANDG") = rowICTITEM1.Item("ITEM_COST_LANDG")
            row.Item("ITEM_COST_TOOLG") = rowICTITEM1.Item("ITEM_COST_TOOLG")
        Next

        With dst.Tables("ASTSRPT1").Columns
            .Add("AMT_BOM_PREV", GetType(System.Decimal), "QTY_BOM * STD_COST_LM")
            .Add("AMT_REVAL", GetType(System.Decimal), "QTY_BOM * (STD_COST_TM - STD_COST_LM)")
            For Each C As String In New String() {"BOM", "SHP", "RTN", "REC", "ADJ", "CON", "EOM"}
                .Add("AMT_" & C, GetType(System.Decimal), "QTY_" & C & " * STD_COST_TM")
            Next
            .Add("AMT_VCOST", GetType(System.Decimal), "QTY_EOM * ISNULL(ITEM_COST_VCOST, 0)")
            .Add("AMT_MATLS", GetType(System.Decimal), "QTY_EOM * (ISNULL(ITEM_COST_MATLS, 0) + ISNULL(ITEM_COST_LANDGI, 0) + ISNULL(ITEM_COST_TOOLGI, 0))")
            .Add("AMT_LANDG", GetType(System.Decimal), "QTY_EOM * ISNULL(ITEM_COST_LANDG, 0)")
            .Add("AMT_TARIFF", GetType(System.Decimal), "QTY_EOM * ISNULL(ITEM_COST_TOOLG, 0)")
        End With

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")


        grdASTEXPT1.Text = "Inventory Valuation by Item - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
        Next
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        If Not Cs.Contains("COLLECTION_CODE") Then Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70)
        If Not Cs.Contains("ITEM_COST_MAKE_BUY") Then Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY", "MB", 30)
        If Not Cs.Contains("COST_CATGY_CODE") Then Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "Cost Catgy", 70)
        If Not Cs.Contains("PROD_CODE") Then Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70)
        If Not Cs.Contains("VEND_CODE") Then Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        Set_DX_Column(grdASTEXPT1, "STD_COST_LM", "Std Prev", 90, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "STD_COST_TM", "Std Curr", 90, "#.0000", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "QTY_BOM", "#Beg", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_SHP", "#Shp", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_RTN", "#Rtn", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_REC", "#Rec", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_ADJ", "#Adj", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_CON", "#Con", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_EOM", "#End", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "AMT_BOM_PREV", "$Beg@Prev", 90, "#,##0", , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "AMT_REVAL", "$ReVal", 90, "#,##0", , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "AMT_BOM", "$Beg@Curr", 90, "#,##0", , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "AMT_SHP", "$Shp", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_RTN", "$Rtn", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_REC", "$Rec", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_ADJ", "$Adj", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_CON", "$Con", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_EOM", "$End", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_VCOST", "$VCost", 90, "#,##0", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "AMT_MATLS", "$Matls", 90, "#,##0", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "AMT_LANDG", "$Landg", 90, "#,##0", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "AMT_TARIFF", "$Tariff", 90, "#,##0", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PV_DEF", "$PPV", 90, "#,##0", , Color.Violet)
        Set_DX_Column(grdASTEXPT1, "MV_DEF", "$MUV", 90, "#,##0", , Color.Violet)
        Set_DX_Column(grdASTEXPT1, "FV_DEF", "$FPV", 90, "#,##0", , Color.Violet)
        Set_DX_Column(grdASTEXPT1, "TV_DEF", "$TPV", 90, "#,##0", , Color.Violet)

        Create_Summary(grdASTEXPT1, "AMT_BOM_PREV")
        Create_Summary(grdASTEXPT1, "AMT_REVAL")

        For Each C As String In New String() _
            {"QTY_BOM", "QTY_SHP", "QTY_RTN", "QTY_REC", "QTY_ADJ", "QTY_CON", "QTY_EOM",
             "AMT_BOM", "AMT_SHP", "AMT_RTN", "AMT_REC", "AMT_ADJ", "AMT_CON", "AMT_EOM", "AMT_VCOST", "AMT_MATLS", "AMT_LANDG", "AMT_TARIFF", "PV_DEF", "MV_DEF", "FV_DEF", "TV_DEF"}
            Create_Summary(grdASTEXPT1, C)
        Next


        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub


    Sub Prepare_Work_File()

        RYP_LM = ASCMAIN1.Period_Calc(RYP, -1)

        Dim sqlICTSTAT1 As String = ", (Select ICTSTAT1.ITEM_CODE,ICTSTAT1.WHSE_CODE" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_BEG,0)) QTY_BOM" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_SHP,0)) QTY_SHP" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_RTN,0)) QTY_RTN" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_REC,0)) QTY_REC" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_ADJ,0)) QTY_ADJ" _
        & ", Sum (NVL(ICTSTAT1.WHSE_QTY_CON,0)) QTY_CON" _
        & " from ICTSTAT1 " _
        & " where ICTSTAT1.OPS_YYYYPP = '" & RYP & "'" _
        & " group by ICTSTAT1.ITEM_CODE,ICTSTAT1.WHSE_CODE) ICTSTAT1"

        Dim sqlICTSTAT5 As String = ", (Select ICTSTAT5.ITEM_CODE,ICTSTAT5.WHSE_CODE" _
        & ", Sum (NVL(ICTSTAT5.WHSE_QTY_ON_HAND,0)) QTY_EOM" _
        & " from " & IIf(RYP = ASCMAIN1.CYP, "ICTSTAT2", "ICTSTAT5") & " ICTSTAT5" _
        & IIf(RYP = ASCMAIN1.CYP, "", " where ICTSTAT5.OPS_YYYYPP = '" & RYP & "'") _
        & " group by ICTSTAT5.ITEM_CODE,ICTSTAT5.WHSE_CODE) ICTSTAT5"

        ASCMAIN1.sql = "Select ICTITEMW.ITEM_CODE, ICTITEMW.WHSE_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_COST_MAKE_BUY" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTITEM1.VEND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
        & ", ICTCOSTA_LM.ITEM_COST_TOTAL STD_COST_LM" & vbCrLf _
        & ", ICTCOSTA_TM.ITEM_COST_TOTAL STD_COST_TM" & vbCrLf _
        & ", ICTSTAT1.QTY_BOM" & vbCrLf _
        & ", ICTSTAT1.QTY_SHP" & vbCrLf _
        & ", ICTSTAT1.QTY_RTN" & vbCrLf _
        & ", ICTSTAT1.QTY_REC" & vbCrLf _
        & ", ICTSTAT1.QTY_ADJ" & vbCrLf _
        & ", ICTSTAT1.QTY_CON" & vbCrLf _
        & ", ICTSTAT5.QTY_EOM" & vbCrLf _
        & ", ICTIVAR1.PV_DEF" & vbCrLf _
        & ", ICTIVAR1.MV_DEF" & vbCrLf _
        & ", ICTIVAR1.FV_DEF" & vbCrLf _
        & ", ICTIVAR1.TV_DEF" & vbCrLf _
        & " from ICTITEM1, ICTIVAR1, ICTCOSTA ICTCOSTA_LM, ICTCOSTA ICTCOSTA_TM, " & vbCrLf _
        & "(Select DISTINCT ITEM_CODE, WHSE_CODE from ICTSTAT1 " & vbCrLf _
        & " where OPS_YYYYPP >= '" & RYP_LM & "' and OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        & " union " & vbCrLf _
        & " Select DISTINCT ITEM_CODE, WHSE_CODE from ICTSTAT5 " & vbCrLf _
        & " where OPS_YYYYPP >= '" & RYP_LM & "' and OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
        & " union " & vbCrLf _
        & " Select DISTINCT ITEM_CODE, WHSE_CODE from ICTSTAT2) ICTITEMW" _
        & sqlICTSTAT1 _
        & sqlICTSTAT5 _
        & " where ICTIVAR1.ITEM_CODE (+) = ICTITEMW.ITEM_CODE" _
        & "   and ICTIVAR1.OPS_YYYYPP (+) = '" & RYP & "'" _
        & "   and ICTCOSTA_LM.OPS_YYYYPP (+) = '" & RYP_LM & "'" _
        & "   and ICTCOSTA_LM.ITEM_CODE (+) = ICTITEMW.ITEM_CODE" _
        & "   and ICTCOSTA_TM.OPS_YYYYPP (+) = '" & RYP & "'" _
        & "   and ICTCOSTA_TM.ITEM_CODE (+) = ICTITEMW.ITEM_CODE" _
        & "   and ICTSTAT1.ITEM_CODE (+) = ICTITEMW.ITEM_CODE" _
        & "   and ICTSTAT1.WHSE_CODE (+) = ICTITEMW.WHSE_CODE" _
        & "   and ICTSTAT5.ITEM_CODE (+) = ICTITEMW.ITEM_CODE" _
        & "   and ICTSTAT5.WHSE_CODE (+) = ICTITEMW.WHSE_CODE" _
        & "   and ICTITEM1.ITEM_CODE = ICTITEMW.ITEM_CODE"
        ICTIVAL1 = ASCMAIN1.Temp_Table

        '   ASCDATA1.ExecuteSQL("Delete from " & ICTIVAL1 & " where NVL(QTY_BOM,0) = 0 and NVL(QTY_EOM,0) = 0")
        If Absx1.chkFor("SUPPRESS_ZERO").Checked Then
            ASCDATA1.ExecuteSQL("Delete from " & ICTIVAL1 & " where NVL(QTY_BOM, 0) = 0 And NVL(QTY_SHP, 0) = 0 And NVL(QTY_RTN, 0) = 0 And NVL(QTY_REC, 0) = 0 And NVL(QTY_ADJ, 0) = 0 And NVL(QTY_CON, 0) = 0 And NVL(QTY_EOM, 0) = 0 And NVL(PV_DEF, 0) = 0 And NVL(MV_DEF, 0) = 0")
        End If

        If Absx1.chkFor("VALUE_EXP_AS_PURCH_AS_ZERO").Checked Then
            ASCDATA1.ExecuteSQL("Update " & ICTIVAL1 & " Set STD_COST_LM = 0, STD_COST_TM = 0 where ITEM_CODE in (Select ITEM_CODE from ICTITEM1 where COST_CATGY_CODE in (Select COST_CATGY_CODE from ICTCOST1 where EXP_AT_PURCHASE = '1'))")
        End If

        ASCMAIN1.sql = "Select * from " & ICTIVAL1
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTIVAL1", 2))

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If
        End If
    End Sub

End Class