Public Class ICRIVAR1

    Dim Expensed_or_Deferred As String
    Dim NYP As String
    Dim FYP As String
    Dim ICTIVAR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 12, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Expensed_or_Deferred = Absx1.optFor("ED").Value
        NYP = ASCMAIN1.Period_Calc(RYP, 1)
        FYP = Mid(RYP, 1, 4) & "01"

        Prepare_dst(True)
    End Sub

    Public Overrides Sub Print_Report()
        If Expensed_or_Deferred = "D" Then
            SUBT = "Showing Variances Deferred in " & RYPLEGEND
        Else
            If optED.Value = "M" Then
                SUBT = "Showing Variances Expensed MTD in " & RYPLEGEND
            Else
                SUBT = "Showing Variances Expensed YTD as of " & RYPLEGEND
            End If
        End If

        CR_params.Add("ED", Expensed_or_Deferred)
        Generate_Report(RPT, , SUBT)

        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("ICTIVAR1")

        grdASTEXPT1.Text = "Inventory Variances by Item " & optED.Text & " - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 100, , , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 70)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY", "MB", 30)
        Set_DX_Column(grdASTEXPT1, "ITEM_BASIC_PROMO", "BP", 30)
        Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "Cost Catgy", 70)
        Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100)

        Set_DX_Column(grdASTEXPT1, "ITEM_COST_STD", "Std Cost", 90, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "WHSE_QTY_ON_HAND", "#On Hand", 90, "#,##0", , Color.Orange)

        If Expensed_or_Deferred = "D" Then
            ' DO NOT SHOW EXP COLS
        Else
            Set_DX_Column(grdASTEXPT1, "PV", "Price $Var", 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "MV", "Matls $Var", 90, "#,##0", , Color.LightGreen)
            '  Set_DX_Column(grdASTEXPT1, "CV", "Curr $Var", 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "RV", "Reval $Var", 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "FV", "Frt $Var", 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "TV", "Trf $Var", 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "TOTAL_V", "PV+MV+FV+TV+RV", 90, "#,##0", , Color.LightGreen)
        End If

        Set_DX_Column(grdASTEXPT1, "PV_DEF", "PPV Def $Var", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "MV_DEF", "MUV Def $Var", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "FV_DEF", "FrV Def $Var", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TV_DEF", "TfV Def $Var", 90, "#,##0", , Color.LightBlue)

        Set_DX_Column(grdASTEXPT1, "TOTAL_VD", "Total $Def", 90, "#,##0", , Color.LightBlue)

        If Expensed_or_Deferred = "D" Then
        Else
            Set_DX_Column(grdASTEXPT1, "PV_DEF_REV", "PPV Def $Rev", 90, "#,##0", , Color.LightPink)
            Set_DX_Column(grdASTEXPT1, "MV_DEF_REV", "MUV Def $Rev", 90, "#,##0", , Color.LightPink)
            Set_DX_Column(grdASTEXPT1, "FV_DEF_REV", "FrV Def $Rev", 90, "#,##0", , Color.LightPink)
            Set_DX_Column(grdASTEXPT1, "TV_DEF_REV", "TfV Def $Rev", 90, "#,##0", , Color.LightPink)
            Set_DX_Column(grdASTEXPT1, "PV_DEF_EXP", "PPV Def $Exp", 90, "#,##0", , Color.LightSeaGreen)
            Set_DX_Column(grdASTEXPT1, "MV_DEF_EXP", "MUV Def $Exp", 90, "#,##0", , Color.LightSeaGreen)
            Set_DX_Column(grdASTEXPT1, "FV_DEF_EXP", "FrV Def $Exp", 90, "#,##0", , Color.LightSeaGreen)
            Set_DX_Column(grdASTEXPT1, "TV_DEF_EXP", "TfV Def $Exp", 90, "#,##0", , Color.LightSeaGreen)

            Set_DX_Column(grdASTEXPT1, "RV_EXP_VCOST", "RV Pur", 90, "#,##0", , Color.Azure)
            Set_DX_Column(grdASTEXPT1, "RV_EXP_LANDG", "RV Frt", 90, "#,##0", , Color.Azure)
            Set_DX_Column(grdASTEXPT1, "RV_EXP_TOOLG", "RV Trf", 90, "#,##0", , Color.Azure)
            'Set_DX_Column(grdASTEXPT1, "RV_EXP_OVRHD", "RV Ovh", 90, "#,##0", , Color.Azure)
            Set_DX_Column(grdASTEXPT1, "RV_EXP_MATLS", "RV Pur_M", 90, "#,##0", , Color.Azure)
            Set_DX_Column(grdASTEXPT1, "RV_EXP_LANDGI", "RV Frt_M", 90, "#,##0", , Color.Azure)
            Set_DX_Column(grdASTEXPT1, "RV_EXP_TOOLGI", "RV Trf_M", 90, "#,##0", , Color.Azure)
            'Set_DX_Column(grdASTEXPT1, "RV_EXP_OVRHDI", "RV Ovh_M", 90, "#,##0", , Color.Azure)

        End If

        If Expensed_or_Deferred = "M" Then
            Set_DX_Column(grdASTEXPT1, "TOTAL_VX", "Total $Var", 90, "#,##0", , Color.Pink) ' WHY WOULD WE EVER WANT TO SHOW EXP + DEF?
        End If

        For Each C As String In New String() _
            {"WHSE_QTY_ON_HAND", "PV", "MV", "CV", "RV", "FV", "TV", "PV_DEF", "MV_DEF", "FV_DEF", "TV_DEF", "TOTAL_V", "TOTAL_VD", "TOTAL_VX"}
            Create_Summary(grdASTEXPT1, C)
        Next
        For Each C As String In New String() _
            {"PV_DEF_REV", "MV_DEF_REV", "FV_DEF_REV", "TV_DEF_REV", "PV_DEF_EXP", "MV_DEF_EXP", "FV_DEF_EXP", "TV_DEF_EXP"}
            Create_Summary(grdASTEXPT1, C)
        Next
        For Each C As String In New String() _
            {"RV_EXP_VCOST", "RV_EXP_LANDG", "RV_EXP_TOOLG"} ' , "RV_EXP_OVRHD"}
            Create_Summary(grdASTEXPT1, C)
        Next
        For Each C As String In New String() _
            {"RV_EXP_MATLS", "RV_EXP_LANDGI", "RV_EXP_TOOLGI"} ' , "RV_EXP_OVRHDI"}
            Create_Summary(grdASTEXPT1, C)
        Next

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub

    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        With dst
            Create_Work_File(True)

            ASCMAIN1.sql = "Select * from " & ICTIVAR1
            Create_TDA(.Tables.Add, "ICTIVAR1", "**", 0, False, , 1)
            .Tables("ICTIVAR1").Columns.Add("TOTAL_V", GetType(System.Decimal), "ISNULL(PV,0)+ISNULL(MV,0)+ISNULL(CV,0)+ISNULL(RV,0)+ISNULL(FV,0)+ISNULL(TV,0)")
            .Tables("ICTIVAR1").Columns.Add("TOTAL_VX", GetType(System.Decimal), "ISNULL(PV,0)+ISNULL(MV,0)+ISNULL(CV,0)+ISNULL(RV,0)+ISNULL(FV,0)+ISNULL(TV,0)+ISNULL(PV_DEF,0)+ISNULL(MV_DEF,0)+ISNULL(FV_DEF,0)+ISNULL(TV_DEF,0)")
            .Tables("ICTIVAR1").Columns.Add("TOTAL_VD", GetType(System.Decimal), "ISNULL(PV_DEF,0)+ISNULL(MV_DEF,0)+ISNULL(FV_DEF,0)+ISNULL(TV_DEF,0)")

            For Each TABLE_NAME As String In New String() {"ICTBRAN1", "ICTCOLL1", "ICTCOST1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, TABLE_NAME, 1))
            Next
        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Loading Data", "")

        EnforceConstraints(False)
        Create_Work_File(False)
        Fill_Records("ICTIVAR1")
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
      
    Overrides Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Sub Verify_Special(eItemKey As String)
        MyBase.Verify_Special(eItemKey)

        Select Case eItemKey
            Case "Proceed"
        End Select
    End Sub

    Sub Create_Work_File(initialize As Boolean)

        Dim sqlw As String = ""
        If initialize Then
            sqlw = " and ROWNUM < 1"
        Else
            sqlw &= SQLA_filter("PROD_CODE", "ICTITEM1")
            sqlw &= SQLA_filter("COLLECTION_CODE", "ICTITEM1")
            sqlw &= SQLA_filter("BRAND_CODE", "ICTCOLL1")
            sqlw &= SQLA_filter("HC_CODE", "ICTCOLL1")
            sqlw &= SQLA_filter("COST_CATGY_CODE", "ICTITEM1")
            sqlw &= SQLA_filter("ITEM_BASIC_PROMO", "ICTITEM1")
        End If

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_MAKE_BUY, ICTITEM1.VEND_CODE" & vbCrLf _
            & " from ICTITEM1, ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & sqlw
        If initialize Then
            ICTIVAR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add WHSE_QTY_ON_HAND NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add PV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add MV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add CV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add FV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add TV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add PV_DEF NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add MV_DEF NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add FV_DEF NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add TV_DEF NUMBER (12,6)")

            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add PV_DEF_REV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add MV_DEF_REV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add FV_DEF_REV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add TV_DEF_REV NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add PV_DEF_EXP NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add MV_DEF_EXP NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add FV_DEF_EXP NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add TV_DEF_EXP NUMBER (12,6)")

            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_VCOST NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_LANDG NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_TOOLG NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_OVRHD NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_MATLS NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_LANDGI NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_TOOLGI NUMBER (12,6)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIVAR1 & " Add RV_EXP_OVRHDI NUMBER (12,6)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTIVAR1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTIVAR1 _
            & " (ITEM_CODE,ITEM_DESC,ITEM_UOM,ITEM_COST_STD,BRAND_CODE,COLLECTION_CODE,COST_CATGY_CODE,ITEM_BASIC_PROMO,PROD_CODE,ITEM_COST_MAKE_BUY,VEND_CODE) " _
            & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ITEM_CODE, Sum (WHSE_QTY_BEG) WHSE_QTY_BEG from ICTSTAT1" & vbCrLf _
                & " where OPS_YYYYPP = '" & NYP & "'" & vbCrLf _
                & "   and ITEM_CODE in (Select ITEM_CODE from " & ICTIVAR1 & ")" & vbCrLf _
                & " group by ITEM_CODE"
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & ICTIVAR1 & vbCrLf _
                & "    Set WHSE_QTY_ON_HAND = R1.WHSE_QTY_BEG" & vbCrLf _
                & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf _
                & ", Sum (PV_EXP) PV" & vbCrLf _
                & ", Sum (MV_EXP) MV" & vbCrLf _
                & ", Sum (CV_EXP) CV" & vbCrLf _
                & ", Sum (RV_EXP) RV" & vbCrLf _
                & ", Sum (FV_EXP) FV" & vbCrLf _
                & ", Sum (TV_EXP) TV" & vbCrLf _
                & ", Sum (PV_DEF) PV_DEF" & vbCrLf _
                & ", Sum (MV_DEF) MV_DEF" & vbCrLf _
                & ", Sum (FV_DEF) FV_DEF" & vbCrLf _
                & ", Sum (TV_DEF) TV_DEF" & vbCrLf _
                & ", Sum (RV_EXP_VCOST) RV_EXP_VCOST" & vbCrLf _
                & ", Sum (RV_EXP_LANDG) RV_EXP_LANDG" & vbCrLf _
                & ", Sum (RV_EXP_TOOLG) RV_EXP_TOOLG" & vbCrLf _
                & ", Sum (RV_EXP_OVRHD) RV_EXP_OVRHD" & vbCrLf _
                & ", Sum (RV_EXP_MATLS) RV_EXP_MATLS" & vbCrLf _
                & ", Sum (RV_EXP_LANDGI) RV_EXP_LANDGI" & vbCrLf _
                & ", Sum (RV_EXP_TOOLGI) RV_EXP_TOOLGI" & vbCrLf _
                & ", Sum (RV_EXP_OVRHDI) RV_EXP_OVRHDI" & vbCrLf _
                & " from ICTIVAR1"
            If Expensed_or_Deferred = "D" Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "_EXP)", "_DEF)") & " where OPS_YYYYPP = '" & RYP & "'"
            Else
                If optED.Value = "M" Then
                    ASCMAIN1.sql &= " where OPS_YYYYPP >= '" & RYP & "' and OPS_YYYYPP <= '" & RYP & "'"
                Else
                    ASCMAIN1.sql &= " where OPS_YYYYPP >= '" & FYP & "' and OPS_YYYYPP <= '" & RYP & "'"
                End If
            End If

            ASCMAIN1.sql &= "" _
                & "   and ITEM_CODE in (Select ITEM_CODE from " & ICTIVAR1 & ")" & vbCrLf _
                & " group by ITEM_CODE"
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & ICTIVAR1 & vbCrLf _
                & "    Set PV = R1.PV, MV = R1.MV, CV = R1.CV, RV = R1.RV, FV = R1.FV, TV = R1.TV" & vbCrLf _
                & "      , PV_DEF = R1.PV_DEF, MV_DEF = R1.MV_DEF, FV_DEF = R1.FV_DEF, TV_DEF = R1.TV_DEF" & vbCrLf _
                & "      , RV_EXP_VCOST = R1.RV_EXP_VCOST, RV_EXP_LANDG = R1.RV_EXP_LANDG, RV_EXP_TOOLG = R1.RV_EXP_TOOLG, RV_EXP_OVRHD = R1.RV_EXP_OVRHD" & vbCrLf _
                & "      , RV_EXP_MATLS = R1.RV_EXP_MATLS, RV_EXP_LANDGI = R1.RV_EXP_LANDGI, RV_EXP_TOOLGI = R1.RV_EXP_TOOLGI, RV_EXP_OVRHDI = R1.RV_EXP_OVRHDI" & vbCrLf _
                & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSQL($"Update {ICTIVAR1} Set PV_DEF_EXP = PV_DEF, MV_DEF_EXP = MV_DEF, FV_DEF_EXP = FV_DEF, TV_DEF_EXP = TV_DEF")


            If Expensed_or_Deferred = "D" Then
            Else

                ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf _
                & ", Sum (-1 * PV_DEF) PV_DEF" & vbCrLf _
                & ", Sum (-1 * MV_DEF) MV_DEF" & vbCrLf _
                & ", Sum (-1 * FV_DEF) FV_DEF" & vbCrLf _
                & ", Sum (-1 * TV_DEF) TV_DEF" & vbCrLf _
                & " from ICTIVAR1"

                If optED.Value = "M" Then
                    ASCMAIN1.sql &= " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -1) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(RYP, -1) & "'"
                Else
                    ASCMAIN1.sql &= " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(FYP, -1) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(RYP, -1) & "'"
                End If

                ASCMAIN1.sql &= "" _
                & "   and ITEM_CODE in (Select ITEM_CODE from " & ICTIVAR1 & ")" & vbCrLf _
                & " group by ITEM_CODE"
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update " & ICTIVAR1 & " Set " & vbCrLf _
                    & "     PV_DEF = NVL(PV_DEF,0) + NVL(R1.PV_DEF,0)" & vbCrLf _
                    & "   , MV_DEF = NVL(MV_DEF,0) + NVL(R1.MV_DEF,0)" & vbCrLf _
                    & "   , FV_DEF = NVL(FV_DEF,0) + NVL(R1.FV_DEF,0)" & vbCrLf _
                    & "   , TV_DEF = NVL(TV_DEF,0) + NVL(R1.TV_DEF,0)" & vbCrLf _
                    & "   , PV_DEF_REV = NVL(R1.PV_DEF,0)" & vbCrLf _
                    & "   , MV_DEF_REV = NVL(R1.MV_DEF,0)" & vbCrLf _
                    & "   , FV_DEF_REV = NVL(R1.FV_DEF,0)" & vbCrLf _
                    & "   , TV_DEF_REV = NVL(R1.TV_DEF,0)" & vbCrLf _
                    & "    where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

            End If


            ASCMAIN1.sql = "Delete from " & ICTIVAR1 & " where NVL(PV,0) = 0 and NVL(MV,0) = 0 and NVL(CV,0) = 0 and NVL(RV,0) = 0 and NVL(FV,0) = 0 and NVL(TV,0) = 0  and NVL(PV_DEF,0) = 0 and NVL(MV_DEF,0) = 0 and NVL(FV_DEF,0) = 0 and NVL(TV_DEF,0) = 0"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub
End Class