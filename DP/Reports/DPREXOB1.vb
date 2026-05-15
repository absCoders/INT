Public Class DPREXOB1
    Dim PRD_END_DATE As String
    Dim NYP As String
    Dim MOS As Integer

    Dim sqlDPTMRPGO As String
    Dim DPTMRPGO As String

    Dim RYP_desc As String
    Dim RYP_end As String

    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 18, 0)
        Me.grpCalcExcessMonth.Visible = chkMAXPOS.Checked
        ' TAC.DPCMAIN1.Create_Worktables_DPTMRPGO(True, Me, sqlDPTMRPGO, DPTMRPGO, RYP_end, False)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        MOS = Val(Absx1.numFor("FPDMOS").Value & "")
        RYP = ASCMAIN1.CYP
        NYP = ASCMAIN1.Period_Calc(RYP, MOS - 1)
        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", NYP)
        PRD_END_DATE = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")

        Dim DTEs(12) As String
        Dim PRDs(12) As String
        For M As Integer = 0 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, M - 1)
            Dim row As DataRow = LookUp("GLTPARM2", YP)
            DTEs(M) = Format(row.Item("PRD_END_DATE"), "dd-MMM-yyyy")
            PRDs(M) = YP
        Next


        ' New FC_MaxPos

        If chkMAXPOS.Checked Then
            RYP_desc = Absx1.cmbFor("RYP").Value
            RYP_end = Mid(RYP_desc, 1, 4) & Mid(RYP_desc, 6, 2)

            TAC.DPCMAIN1.Create_Worktables_DPTMRPGO(True, Me, sqlDPTMRPGO, DPTMRPGO, RYP_end, False)

            'With dst

            '    ASCMAIN1.sql = $"Select * from {DPTMRPGO}"
            '    Create_TDA(.Tables.Add, "DPTMRPGO", "**", 0, False, "", 2)
            '    With .Tables("DPTMRPGO").Columns
            '        .Add("OVER_QTY", GetType(System.Int64), "ISNULL(EOM,0)-ISNULL(FC,0)")
            '        .Add("OVER_EXT_COST", GetType(System.Decimal), "ISNULL(OVER_QTY,0)*ISNULL(ITEM_COST_STD,0)")
            '        .Add("FC_CUR", GetType(System.Int64))
            '        .Add("FC_FUT", GetType(System.Int64))
            '        .Add("PO_CUR", GetType(System.Int64))
            '        .Add("PO_FUT", GetType(System.Int64))
            '        .Add("ZERO", GetType(System.String), "IIF(ISNULL(EOM,0)=0 AND ISNULL(FC_CUR,0)=0 AND ISNULL(FC_FUT,0)=0 AND ISNULL(PO_CUR,0)=0 AND ISNULL(PO_FUT,0)=0 AND ISNULL(OVER_QTY,0)=0 AND ISNULL(FC,0)=0,'0','1')")
            '    End With

            'End With

            'dst.Tables("DPTMRPGO").Rows.Clear()

            TAC.DPCMAIN1.Create_Worktables_DPTMRPGO(False, Me, sqlDPTMRPGO, DPTMRPGO, RYP_end, False)
        End If

        ' Page0

        Page0.Add("Number of Months of Demand: " & CStr(MOS))
        Page0.Add("Showing Excess or Obsolete: " & Absx1.optFor("OPTXOB").Text)
        Page0.Add("Showing Items with Excess/Obsolete: " & Absx1.optFor("OPTHOB").Text)
        Page0.Add("Calculating Demand from: " & Absx1.optFor("OPTDEMAND").Text)
        If Absx1.chkFor("CHKPLANS").Checked Then Page0.Add("Including Plans Supply Data")
        If Absx1.chkFor("CHKPLAND").Checked Then Page0.Add("Including Plans Demand Data")
        If Absx1.chkFor("CHKSORT").Checked Then Page0.Add("Sorting Items by Excess/Obsolete Value")
        If Absx1.chkFor("CHKEXCLNONMRP").Checked Then Page0.Add("Excluding Supply & Demand in non-MRP Warehouses")
        If Absx1.chkFor("CHKSINGLEWHSE").Checked Then Page0.Add("Only Warehouse: " & Absx1.txtFor("WHSE_CODE").Text)
        If Absx1.optFor("OPTDEMAND").Value = "F" And Absx1.chkFor("OVRSHP").Checked Then
            Page0.Add(Absx1.chkFor("OVRSHP").Text)
        End If
        ' Prepare Work Tables

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""
        If Absx1.chkFor("CHKEXCLNONMRP").Checked Then
            sql_filter &= " and (NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1')"
        End If
        If Absx1.chkFor("CHKSINGLEWHSE").Checked Then
            sql_filter &= " and ICTWHSE1.WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End If

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        Dim sql_filter2 As String = ""
        Dim sql_Having As String = ""

        ASCMAIN1.Progress("Status")
        MyBase.Get_SQL("S")
        Get_Data("ICTSTAT2", _
                    sql_Sum:=", Sum (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) QTY_ON_HAND", _
                    sql_Sum_Cols:="QTY_ON_HAND", _
                    sql_filter:=sql_filter, _
                    sql_filter2:="", _
                    sql_Having:="Sum (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) <> 0", _
                    sql_Appended_Cols:=", ICTSTAT2.ITEM_CODE")

        ASCMAIN1.Progress("Purchase Orders")
        MyBase.Get_SQL("O")
        Get_Data("POTORDR2", _
            sql_Sum:=", Sum (NVL(POTORDR2.PO_QTY_OPN,0)) QTY_ON_ORDER", _
            sql_Sum_Cols:="QTY_ON_ORDER", _
            sql_filter:=sql_filter, _
            sql_filter2:=" and POTORDR2.PO_DATE_REQUIRED <= '" & PRD_END_DATE & "'", _
            sql_Having:="Sum (NVL(POTORDR2.PO_QTY_OPN,0)) <> 0", _
            sql_Appended_Cols:=", POTORDR2.ITEM_CODE")

        ASCMAIN1.Progress("Plans")
        MyBase.Get_SQL("P")
        Get_Data("DPTPLAN1", _
            sql_Sum:=", Sum (NVL(DPTPLAN1.QTY_PLANNED,0)) QTY_PLANNED", _
            sql_Sum_Cols:="QTY_PLANNED", _
            sql_filter:=sql_filter, _
            sql_filter2:=" and DPTPLAN1.DATE_REQUIRED <= '" & PRD_END_DATE & "'", _
            sql_Having:="Sum (NVL(DPTPLAN1.QTY_PLANNED,0)) <> 0", _
            sql_Appended_Cols:=", DPTPLAN1.ITEM_CODE")

        ASCMAIN1.Progress("Activity")
        MyBase.Get_SQL("A")
        Get_Data("ICTSTAT1", _
            sql_Sum:=", Sum (ICTSTAT1.WHSE_QTY_SHP) MTD_SHP, Sum (ICTSTAT1.WHSE_QTY_RTN) MTD_RTN", _
            sql_Sum_Cols:="MTD_SHP,MTD_RTN", _
            sql_filter:=sql_filter, _
            sql_filter2:=" and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'", _
            sql_Having:="Sum (ICTSTAT1.WHSE_QTY_SHP) <> 0 or Sum (ICTSTAT1.WHSE_QTY_RTN) <> 0", _
            sql_Appended_Cols:=", ICTSTAT1.ITEM_CODE")

        ASCMAIN1.Progress("Sales Orders")
        Dim sqlOPs As String = ""
        For M As Integer = 0 To 3
            Dim MLAST As Integer = 0
            If M > 0 Then MLAST = M - 1
            sqlOPs &= ", Sum (Case When " & IIf(M = 0, "", "SOTORDR1.ORDR_SHIP_DATE > '" & DTEs(MLAST) & "' and ") & "SOTORDR1.ORDR_SHIP_DATE <= '" & DTEs(M) & "' Then NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTYOP_" & Format(M, "00")
        Next
        MyBase.Get_SQL("L")
        Get_Data("SOTORDR2", _
            sql_Sum:=", Sum (SOTORDR2.ORDR_QTY_OPEN) QTY_OPEN_SO, Sum (SOTORDR2.ORDR_QTY_PICK) QTY_PICK_SO" & sqlOPs, _
            sql_Sum_Cols:="QTY_OPEN_SO,QTY_PICK_SO,QTYOP_00,QTYOP_01,QTYOP_02,QTYOP_03", _
            sql_filter:=sql_filter, _
            sql_filter2:=" and SOTORDR2.ORDR_STATUS in ('O','P')", _
            sql_Having:="Sum (SOTORDR2.ORDR_QTY_OPEN) <> 0 or Sum (SOTORDR2.ORDR_QTY_PICK) <> 0", _
            sql_Appended_Cols:=", SOTORDR2.ITEM_CODE")

        ASCMAIN1.Progress("Purchase Order Committments")
        If Absx1.optFor("OPTDEMAND").Value = "F" Then
            MyBase.Get_SQL("C")
            Get_Data("POTORDR9", _
                sql_Sum:=", Sum (DECODE(POTORDR9.PO_ORDER_LNO,0,0,NVL(POTORDR9.PO_QTY_COM,0))) PROD_COM" _
                       & ", Sum (DECODE(POTORDR9.PO_ORDER_LNO,0,NVL(POTORDR9.PO_QTY_COM,0),0)) PLAN_COM", _
                sql_Sum_Cols:="PROD_COM,PLAN_COM", _
                sql_filter:=sql_filter, _
                sql_filter2:=" and NVL(POTORDR2.PO_DATE_COMPSDUE,DPTPLAN1.DATE_COMPSDUE) <= '" & PRD_END_DATE & "'", _
                sql_Having:="Sum (NVL(POTORDR9.PO_QTY_COM,0)) <> 0", _
                sql_Appended_Cols:=", POTORDR9.ITEM_CODE")
        Else
            Dim sql_filter_history As String = ""
            If Absx1.optFor("OPTDEMAND").Value = "R" Then
                sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * MOS) & "'"
            Else
                sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + MOS) & "'"
            End If

            MyBase.Get_SQL("A")
            Get_Data("ICTSTAT1", _
                sql_Sum:=", Sum (ICTSTAT1.WHSE_QTY_CON) PROD_COM", _
                sql_Sum_Cols:="PROD_COM", _
                sql_filter:=sql_filter, _
                sql_filter2:=sql_filter_history, _
                sql_Having:="Sum (ICTSTAT1.WHSE_QTY_CON) <> 0", _
                sql_Appended_Cols:=", ICTSTAT1.ITEM_CODE")
        End If


        ASCMAIN1.Progress("Demand (" & Absx1.optFor("OPTDEMAND").Text & ")")
        MyBase.Get_SQL(Absx1.optFor("OPTDEMAND").Value)

        Select Case Absx1.optFor("OPTDEMAND").Value
            Case "F"
                Dim sqlFCs As String = ""
                For M As Integer = 0 To 3
                    sqlFCs &= ", Sum (Case When " & IIf(M = 0, "DPTITMF1.OPS_YYYYPP_FC = '000000'", "DPTITMF1.OPS_YYYYPP_FC ='" & PRDs(M) & "'") & " Then NVL(DPTITMF1.FORECAST,0) ELSE 0 END) QTYFC_" & Format(M, "00")
                Next
                Get_Data("DPTITMF1", _
                    sql_Sum:=", Sum (NVL(DPTITMF1.FORECAST,0)) FORECAST" & sqlFCs, _
                    sql_Sum_Cols:="FORECAST,QTYFC_00,QTYFC_01,QTYFC_02,QTYFC_03", _
                    sql_filter:=sql_filter, _
                    sql_filter2:=" and DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and DPTITMF1.OPS_YYYYPP_FC <= '" & NYP & "'", _
                    sql_Having:="Sum (NVL(DPTITMF1.FORECAST,0)) <> 0", _
                    sql_Appended_Cols:=", DPTITMF1.ITEM_CODE")

            Case "R", "H"

                Dim sql_filter_history As String = ""
                If Absx1.optFor("OPTDEMAND").Value = "R" Then
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.CYP & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * MOS) & "'"
                Else
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + MOS) & "'"
                End If

                Get_Data("SOTINVH2", _
                    sql_Sum:=", Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) FORECAST", _
                    sql_Sum_Cols:="FORECAST", _
                    sql_filter:=sql_filter, _
                    sql_filter2:=sql_filter_history, _
                    sql_Having:="Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) <> 0", _
                    sql_Appended_Cols:=", SOTINVH2.ITEM_CODE")
        End Select

        Dim LEVEL As Integer = Get_Level_for("EXC_OBS_IND")

        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        '    ASCMAIN1.sql = "Delete from " & ASTSRPT1 & " where ITEM_CODE <> 'CH011A01' AND ITEM_CODE <> 'CH011B13SF' AND ITEM_CODE <> 'VCA001'"
        '    ASCDATA1.ExecuteSQL()
        'End If

        If LEVEL <> 0 Then
            Dim GX As String = "G" & CStr(LEVEL)
            ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set " & GX & " = 'O'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set " & GX & " = 'E'" _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & ASTSRPT1 _
                & " where FORECAST <> 0 or PROD_COM <> 0 or PLAN_COM <> 0)"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTITEM1.ITEM_SAFETY_STOCK" & vbCrLf _
            & " from ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & ASTSRPT1 & ")"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
        Fill_Records("ICTITEM1")
    End Sub

    Public Overrides Sub Print_Report()

        If Absx1.optFor("OPTDEMAND").Value = "F" Then
            SUBT = "Calculated using " & MOS & " Future Months of Supply & Demand Data (thru " & PRD_END_DATE & ")"
        Else
            SUBT = "Calculated using " & MOS & " Months of Sales History from "
            If Absx1.optFor("OPTDEMAND").Value = "R" Then
                SUBT &= ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * MOS))
            Else
                SUBT &= ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12))
            End If
        End If
        If Absx1.chkFor("CHKEXCLNONMRP").Checked Then
            SUBT &= ", MRP Whses Only"
        End If
        If Absx1.chkFor("CHKSINGLEWHSE").Checked Then
            SUBT &= ", Whse " & Absx1.txtFor("WHSE_CODE").Text & " Only"
        End If

        CR_params.Add("FPDMOS", MOS)

        Generate_Report(RPT, , SUBT)
        Prepare_Data_Extracts()
        If chkMAXPOS.Checked Then
            Prepare_Data_Extracts_FC_MAXPOS()
        Else
            If tabDataExports.Tabs.Count > 1 Then '
                tabDataExports.Tabs(1).Visible = True
            End If
        End If

    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand


        'grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        'grdASTEXPT1.DataSource = dst.Tables("GLTDETL1")

        'grdASTEXPT1.Text = "Account Detail Report"
        'UltraTabControl1.Tabs("Data Exports").Visible = True
        ''  UltraTabControl1.Tabs("Data Grids").Visible = True
        'tabDataExports.Tabs(0).Text = grdASTEXPT1.Text


        With dst.Tables("ASTSRPT1").Columns
            .Add("ITEM_DESC")
            .Add("ITEM_SAFETY_STOCK", GetType(System.Int32))
            .Add("ITEM_RETAIL_PRICE", GetType(System.Decimal))
            '.Add("ITEM_COST_STD", GetType(System.Decimal))
        End With

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            row.Item("ITEM_SAFETY_STOCK") = rowICTITEM1.Item("ITEM_SAFETY_STOCK")
            row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
            'row.Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD")
        Next

        'With dst.Tables("ASTSRPT1").Columns
        '    .Add("AMT_BOM_PREV", GetType(System.Decimal), "QTY_BOM * STD_COST_LM")
        '    .Add("AMT_REVAL", GetType(System.Decimal), "QTY_BOM * (STD_COST_TM - STD_COST_LM)")
        '    For Each C As String In New String() {"BOM", "SHP", "RTN", "REC", "ADJ", "CON", "EOM"}
        '        .Add("AMT_" & C, GetType(System.Decimal), "QTY_" & C & " * STD_COST_TM")
        '    Next
        'End With

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")


        grdASTEXPT1.Text = "Excess / Obsolete Inventory Report - " & SUBT
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

        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        Set_DX_Column(grdASTEXPT1, "ITEM_RETAIL_PRICE", "Retail", 90, "#,##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_STD", "Std Cost", 90, "#.0000", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "QTY_ON_HAND", "Current On Hand", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "ITEM_SAFETY_STOCK", "Safety Stock", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "MTD_SHP", " MTD Shipped", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "FORECAST", "Forecast", 90, "#,##0", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "PROD_COM", "PO Com", 90, "#,##0", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "PLAN_COM", "Pln Com", 90, "#,##0", , Color.LightPink)
        Set_DX_Column(grdASTEXPT1, "QTY_OPEN_SO", "Qty Open", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "QTY_PICK_SO", "Qty Pick", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXC_OH_QTY", "Excess Qty OH", 90, "#,##0", , Color.MediumPurple)
        Set_DX_Column(grdASTEXPT1, "EXC_OH_AMT", "Excess Val OH", 90, "#,##0", , Color.MediumPurple)
        Set_DX_Column(grdASTEXPT1, "QTY_ON_ORDER", "On PO", 90, "#,##0", , Color.Gold)
        Set_DX_Column(grdASTEXPT1, "QTY_PLANNED", "Plans", 90, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "EXC_OO_QTY", "Excess Qty OO", 90, "#,##0", , Color.MediumPurple)
        Set_DX_Column(grdASTEXPT1, "EXC_OO_AMT", "Excess Val OO", 90, "#,##0", , Color.MediumPurple)

        'Set_DX_Column(grdASTEXPT1, "EXC_OO_QTY", "Excess Qty Total", 90, "#,##0", , Color.MediumPurple)
        'Set_DX_Column(grdASTEXPT1, "EXC_OO_AMT", "Excess Val Total", 90, "#,##0", , Color.MediumPurple)

        Create_Summary(grdASTEXPT1, "EXC_OH_AMT")
        Create_Summary(grdASTEXPT1, "EXC_OO_AMT")

        'For Each C As String In New String() {"MV_DEF"}
        '    Create_Summary(grdASTEXPT1, C)
        'Next


        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub

    Sub Prepare_Data_Extracts_FC_MAXPOS()

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
                tabDataExports.Tabs(1).Text = .Text

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
        End With

        'ASCMAIN1.sql = $"Select * from {DPTMRPGO}"
        'Dim DT As DataTable = ASCDATA1.GetDataTable
        grdASTEXPT2.DataSource = dst.Tables("DPTMRPGO")
        Dim dvw As DataView = DirectCast(grdASTEXPT2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "OVER_QTY > 0"

        ASCMAIN1.grdInitializeLayout(grdASTEXPT2)

        Sort_grdColumns(grdASTEXPT2, "ITEM_CODE")
        grdASTEXPT2.Text = $"Items with EOM On Hand over Max Position in {Absx1.cmbFor("RYP").Text} (Excluding Items with All 0's)"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(1).Visible = True
        tabDataExports.Tabs(1).Text = grdASTEXPT2.Text

        Set_DX_Column(grdASTEXPT2, "")

        Set_DX_Column(grdASTEXPT2, "ITEM_CODE", "Item Code", 80)
        Set_DX_Column(grdASTEXPT2, "ITEM_DESC", "Item Description", 150)
        Set_DX_Column(grdASTEXPT2, "ITEM_COST_STD", "Std Cost", 80, "#,##0.00")
        Set_DX_Column(grdASTEXPT2, "BRAND_CODE", "Brand", 80)
        Set_DX_Column(grdASTEXPT2, "COLLECTION_CODE", "Collection", 80)
        Set_DX_Column(grdASTEXPT2, "VEND_CODE", "Vendor", 80)
        Set_DX_Column(grdASTEXPT2, "ITEM_STATUS", "Status", 80)
        Set_DX_Column(grdASTEXPT2, "ITEM_SNU_CODE", "SNU", 40)
        Set_DX_Column(grdASTEXPT2, "ITEM_BASIC_PROMO", "BP", 40)
        Set_DX_Column(grdASTEXPT2, "PROD_CODE", "Product", 80)
        Set_DX_Column(grdASTEXPT2, "ITEM_POS_MAX", "Max Pos", 60, "##0.0")
        Set_DX_Column(grdASTEXPT2, "ITEM_POS_MIN", "Min Pos", 60, "##0.0")
        Set_DX_Column(grdASTEXPT2, "FC", "#FC MaxPos", 80, "#,##0")
        Set_DX_Column(grdASTEXPT2, "EOM", "Qty OH EOM", 80, "#,##0")
        Set_DX_Column(grdASTEXPT2, "EOM_EXT_COST", "$Value EOM", 80, "#,##0", "Sum")
        Set_DX_Column(grdASTEXPT2, "POS", "Act Pos", 60, "##0.00")
        Set_DX_Column(grdASTEXPT2, "OVER_QTY", "Qty Over", 80, "#,##0", "Sum", Color.LightPink)
        Set_DX_Column(grdASTEXPT2, "OVER_EXT_COST", "Amt Over", 80, "#,##0", "Sum", Color.LightPink)
        Set_DX_Column(grdASTEXPT2, "FC_CUR", "FC Cur", 80, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "FC_FUT", "FC Fut", 80, "#,##0", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT2, "PO_CUR", "PO Cur", 80, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT2, "PO_FUT", "PO Fut", 80, "#,##0", , Color.LightBlue)

        'Set_DX_Column(grdASTEXPT2, "QTY_BEG", "#Beg", 80, "#,##0", "Sum", Color.LimeGreen)
        'Set_DX_Column(grdASTEXPT2, "AMT_BEG", "$Beg", 80, "#,##0", "Sum", Color.LimeGreen)
        'Set_DX_Column(grdASTEXPT2, "QTY_REC", "#Rec", 80, "#,##0", "Sum", Color.LightSeaGreen)
        'Set_DX_Column(grdASTEXPT2, "AMT_REC", "$Rec", 80, "#,##0", "Sum", Color.LightSeaGreen)
        'Set_DX_Column(grdASTEXPT2, "QTY_RTN", "#Rtn", 80, "#,##0", "Sum", Color.Gold)
        'Set_DX_Column(grdASTEXPT2, "AMT_RTN", "$Rtn", 80, "#,##0", "Sum", Color.Gold)
        'Set_DX_Column(grdASTEXPT2, "QTY_ADJ", "#Adj", 80, "#,##0", "Sum", Color.Orange)
        'Set_DX_Column(grdASTEXPT2, "AMT_ADJ", "$Adj", 80, "#,##0", "Sum", Color.Orange)

        grdASTEXPT2.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT2, "ITEM_CODE")
        If UltraCheckEditor7.Checked Then
            Sort_grdColumns(grdASTEXPT2, "over_ext_cost")
        End If

        'If ASCMAIN1.CLIENT = "INT" Then
        '    WorkbookView1.GetLock()
        '    Load_DataTable_into_SGXLS(1, 1, DT, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "INV_TYPE,INV_NO", "")
        '    WorkbookView1.ReleaseLock()

        '    WorkbookView1.Visible = True
        '    btnExcel.Visible = True
        '    btnExcel.BringToFront()
        'End If

    End Sub





    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If Absx1.cmbFor("RYP").Value & "" = "" Then
            '    EMsg &= vbCr & "You must Specify a Reporting Period"
            'End If

            If chkSINGLEWHSE.Checked Then
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", txtWHSE_CODE.Text)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "You must Specify a valid Warehouse when choosing a Single Warehouse Report"
                End If
            End If
        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        With dst.Tables("ASTSRPT1")
            If Absx1.optFor("OPTDEMAND").Value = "F" And Absx1.chkFor("OVRSHP").Checked Then
                .Columns("QTY_OVR").Expression = "ISNULL(MTD_SHP,0)+ISNULL(QTYOP_00,0)+ISNULL(QTYOP_01,0)-(ISNULL(QTYFC_00,0)+ISNULL(QTYFC_01,0))"
            Else
                .Columns("QTY_OVR").Expression = "0"
            End If
            .Columns("FC_REM").Expression = "ISNULL(FORECAST,0)+IIF(QTY_OVR>0,QTY_OVR,0)-ISNULL(MTD_SHP,0)"
            .Columns("DEMAND").Expression = "IIF(FC_REM>0,FC_REM,0)+ISNULL(QTY_OPEN_SO,0)+ISNULL(PROD_COM,0)+ISNULL(PLAN_COM,0)"
            .Columns("CUR_AVA").Expression = "ISNULL(QTY_ON_HAND,0)-DEMAND"
            .Columns("PO_AND_PL").Expression = "ISNULL(QTY_ON_ORDER,0)+ISNULL(QTY_PLANNED,0)"
            .Columns("FUT_AVA").Expression = "CUR_AVA+PO_AND_PL"
            .Columns("EXC_OH_QTY").Expression = "IIF(CUR_AVA>0,CUR_AVA,0)"
            .Columns("EXC_OO_QTY").Expression = "IIF(PO_AND_PL>0 AND FUT_AVA>0,FUT_AVA,0)"
            Create_Relation("ICTITEM1", "ASTSRPT1", "ITEM_CODE")
            .Columns.Add("ITEM_COST_STD", GetType(System.Decimal), "PARENT(ICTITEM1_ASTSRPT1).ITEM_COST_STD")
            .Columns("EXC_OH_AMT").Expression = "EXC_OH_QTY*ISNULL(ITEM_COST_STD,0)"
            .Columns("EXC_OO_AMT").Expression = "EXC_OO_QTY*ISNULL(ITEM_COST_STD,0)"
            .Columns.Add("EXC_OBS_IND", GetType(System.String), "IIF(DEMAND=0,'O','X')")
        End With

        Dim exp As String = "0"
        If Absx1.chkFor("CHKSORT").Checked Then
            If Absx1.optFor("OPTHOB").Value = "H" Then
                exp = "EXC_OH_AMT"
            Else
                If Absx1.optFor("OPTHOB").Value = "O" Then
                    exp = "EXC_OO_AMT"
                Else
                    exp = "EXC_OH_AMT + EXC_OO_AMT"
                End If
            End If
        End If
        dst.Tables("ASTSRPT1").Columns.Add("SORT_BY", GetType(System.Decimal), exp)

        Dim sqld As String = ""
        If Absx1.optFor("OPTHOB").Value = "H" Then sqld &= " or EXC_OH_QTY <=0"
        If Absx1.optFor("OPTHOB").Value = "O" Then sqld &= " or EXC_OO_QTY <=0"
        sqld &= " or (EXC_OH_QTY <=0 and EXC_OO_QTY <=0)"

        If Absx1.optFor("OPTXOB").Value = "X" Then sqld &= " or DEMAND = 0"
        If Absx1.optFor("OPTXOB").Value = "O" Then sqld &= " or DEMAND <> 0"
        If sqld <> "" Then
            ASCDATA1.DeleteRows(dst.Tables("ASTSRPT1"), Mid(sqld, 5))
        End If
    End Sub
     
    Private Sub chkSINGLEWHSE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSINGLEWHSE.CheckedChanged
        txtWHSE_CODE.Visible = chkSINGLEWHSE.Checked
    End Sub
     
    Private Sub optDEMAND_ValueChanged(sender As Object, e As EventArgs) Handles optDEMAND.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Absx1.chkFor("OVRSHP").Visible = (optDEMAND.Value = "F")
    End Sub

    Private Sub chkMAXPOS_CheckedChanged(sender As Object, e As EventArgs) Handles chkMAXPOS.CheckedChanged
        lblFPDMOS.Visible = Not chkMAXPOS.Checked
        lblFPDMOS1.Visible = Not chkMAXPOS.Checked
        numFPDMOS.Visible = Not chkMAXPOS.Checked
        Me.grpCalcExcessMonth.Visible = chkMAXPOS.Checked
    End Sub
End Class