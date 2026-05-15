Public Class ICRISTA1

    Private CDCaption As String = String.Empty
    Private optTD As String = String.Empty
    Private optTL As String = String.Empty
    Private optUnitCost As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        ' Load_Popup_Menu(grdICTISTAX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
            '    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
            '    If rowSOTORDR1 IsNot Nothing Then
            '        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
            '    End If

        End Select
    End Sub

#End Region

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)

    End Sub

    Public Overrides Sub Post_Process_Special()
        'grdICTISTAX.DisplayLayout.Bands(0).SortedColumns.Clear()

        'If Not dst.Tables.Contains("ICTISTAX") Then
        '    Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy
        '    tbl.TableName = "ICTISTAX"
        '    dst.Tables.Add(tbl)
        '    Create_Relation("ICTITEM1", "ICTISTAX", "ITEM_CODE")
        '    With dst.Tables("ICTISTAX").Columns
        '        .Add("ITEM_DESC", GetType(System.String), "PARENT.ITEM_DESC")
        '        .Add("ITEM_COST_STD", GetType(System.Decimal), "PARENT.ITEM_COST_STD")
        '        For Each CQTY As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "CON", "RTV", "PHY", "ON_HAND", "ONPO", "OPEN", "PICK", "COMM"}
        '            Dim C As String = "WHSE_QTY_" & CQTY
        '            Dim CC As String = Replace(C, "QTY", "CST")
        '            .Add(CC, GetType(System.Decimal), "ISNULL(" & C & ",0) * ITEM_COST_STD")
        '        Next
        '    End With
        'Else
        '    dst.Tables("ICTISTAX").Merge(dst.Tables("ASTSRPT1"))
        'End If

        'If SEQs > 0 Then
        '    For Each row As DataRow In dst.Tables("ICTISTAX").Select("")
        '        For I As Integer = 1 To SEQs '  COLUMN_CAPTION_by_Lvl.Count - 1
        '            Dim G As String = "G" & CStr(I)
        '            Dim V As String = Split(row.Item(G) & "", ":")(1)
        '            row.Item(G) = V
        '        Next
        '    Next
        'End If

        'grdICTISTAX.DataSource = dst.Tables("ICTISTAX")
        'With grdICTISTAX.DisplayLayout.Bands(0)
        '    With .Columns("ITEM_CODE")
        '        .Header.Caption = "Item Code"
        '        .Width = 100
        '        .Header.Appearance.BackColor = Color.White
        '        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        .Header.Appearance.BackColor2 = Color.LightGray
        '    End With
        '    With .Columns("ITEM_DESC")
        '        .Header.Caption = "Description"
        '        .Header.Appearance.BackColor = Color.White
        '        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        .Header.Appearance.BackColor2 = Color.LightGray
        '    End With
        '    With .Columns("ITEM_COST_STD")
        '        .Header.Caption = "Std Cost"
        '        .Format = "#.0000"
        '        .Header.Appearance.BackColor = Color.White
        '        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        .Header.Appearance.BackColor2 = Color.LightGray
        '    End With
        '    For Each C As String In New String() {"RTV", "PHY", "ONPO", "OPEN", "PICK", "COMM"}
        '        .Columns("WHSE_QTY_" & C).Hidden = True
        '        .Columns("WHSE_CST_" & C).Hidden = True
        '    Next

        '    For I As Integer = 1 To 9
        '        With .Columns("G" & CStr(I))
        '            .Header.Fixed = True
        '            .Header.Appearance.BackColor = Color.White
        '            .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '            .Header.Appearance.BackColor2 = Color.LightGray
        '            If I <= SEQs Then ' COLUMN_CAPTION_by_Lvl.Count - 1 Then
        '                .Header.Caption = COLUMN_CAPTION_by_Lvl(I)
        '                .Hidden = False
        '                .Width = 80
        '            Else
        '                .Hidden = True
        '            End If
        '        End With
        '    Next
        '    .Columns("ITEM_CODE").Header.Fixed = True
        '    .Columns("ITEM_DESC").Header.Fixed = True
        '    .Columns("ITEM_COST_STD").Header.Fixed = True
        'End With

        'grdICTISTAX.DisplayLayout.Bands(0).Summaries.Clear()
        'For Each C As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "CON", "RTV", "PHY", "ON_HAND", "ONPO", "OPEN", "PICK", "COMM"}
        '    With grdICTISTAX.DisplayLayout.Bands(0)
        '        For Each CTYP As String In New String() {"QTY", "CST"}
        '            With .Columns("WHSE_" & CTYP & "_" & C)
        '                .Width = 80
        '                .Format = "#,##0"
        '                .Header.Caption = CTYP & " " & C
        '                .Header.Appearance.BackColor = Color.White
        '                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '                If CTYP = "QTY" Then
        '                    .Header.Appearance.BackColor2 = Color.LightBlue
        '                Else
        '                    .Header.Appearance.BackColor2 = Color.LightGreen
        '                End If
        '            End With
        '            Create_Summary(grdICTISTAX, "WHSE_" & CTYP & "_" & C)
        '        Next
        '    End With
        'Next

        'Sort_grdColumns(grdICTISTAX, "G1, G2, G3, G4, G5, G6, G7, G8, G9, ITEM_CODE")
        ''Show_Filter(grdSOTORDRC, True)
        ''grdSOTORDRC.DisplayLayout.GroupByBox.Hidden = False
        ''grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
        ''grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_CUST_PO", False, True)
        ''   grdICTISTAX.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = True
        'grdICTISTAX.Visible = (optUC.Value = "U")
        '' btnExcelExport.Visible = True

    End Sub
    Public Overrides Sub Print_Report()

        CR_params.Add("SUPPRESS_DTL", IIf(chkSUPPRESS.Checked, "1", "0"))
        CR_params.Add("CD", CDcaption)
        CR_params.Add("NEG", IIf(chkNEG.Checked, "1", "0"))
        CR_params.Add("NEGAVA", IIf(chkNegAvail.Checked, "1", "0"))
        CR_params.Add("ALL_ITEMS", "0")

        If RYP = RYP1 Then
            SUBT = "Status & Activity for " & RYPLEGEND
        Else
            SUBT = "Status & Activity from " & RYPLEGEND & " thru " & RYPLEGEND1
        End If

        Generate_Report("ICRISTA1", String.Empty, SUBT)

        Prepare_Data_Extracts()

    End Sub

    Sub Prepare_Data_Extracts()

        If Not dst.Tables.Contains("ICTISTAX") Then
            Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy
            tbl.TableName = "ICTISTAX"
            dst.Tables.Add(tbl)
            Create_Relation("ICTITEM1", "ICTISTAX", "ITEM_CODE")
            With dst.Tables("ICTISTAX").Columns
                .Add("ITEM_DESC", GetType(System.String), "PARENT.ITEM_DESC")
                .Add("ITEM_COST_STD", GetType(System.Decimal), "PARENT.ITEM_COST_STD")
                For Each CQTY As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "CON", "RTV", "PHY", "ON_HAND", "ONPO", "OPEN", "PICK", "COMM"}
                    Dim C As String = "WHSE_QTY_" & CQTY
                    Dim CC As String = Replace(C, "QTY", "CST")
                    .Add(CC, GetType(System.Decimal), "ISNULL(" & C & ",0) * ITEM_COST_STD")
                Next
                .Add("QTY_AVA", GetType(System.Int64), "ISNULL(WHSE_QTY_ON_HAND,0) - ISNULL(WHSE_QTY_PICK,0) - ISNULL(WHSE_QTY_OPEN,0) - ISNULL(WHSE_QTY_COMM,0)")
                '{ASTSRPT1.WHSE_QTY_ON_HAND} - {ASTSRPT1.WHSE_QTY_PICK} - {ASTSRPT1.WHSE_QTY_OPEN} - {ASTSRPT1.WHSE_QTY_COMM}
                .Add("CST_AVA", GetType(System.Decimal), "ISNULL(QTY_AVA,0) * ITEM_COST_STD")
            End With
        Else
            dst.Tables("ICTISTAX").Merge(dst.Tables("ASTSRPT1"))
        End If

        If SEQs > 0 Then
            For Each row As DataRow In dst.Tables("ICTISTAX").Select("")
                For I As Integer = 1 To SEQs '  COLUMN_CAPTION_by_Lvl.Count - 1
                    Dim G As String = "G" & CStr(I)
                    Dim V As String = Split(row.Item(G) & "", ":")(1)
                    row.Item(G) = V
                Next
            Next
        End If



        grdASTEXPT1.DataSource = dst.Tables("ICTISTAX")
        grdASTEXPT1.Text = "Inventory Roll Forward"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 100)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        Set_DX_Column(grdASTEXPT1, "ITEM_COST_STD", "Std Cost", 90, "#.0000")

        With grdASTEXPT1.DisplayLayout.Bands(0)
            For I As Integer = 1 To 9
                With .Columns("G" & CStr(I))
                    .Header.Fixed = True
                    If I <= SEQs Then ' COLUMN_CAPTION_by_Lvl.Count - 1 Then
                        .Header.Caption = COLUMN_CAPTIONs(I - 1) ' COLUMN_CAPTION_by_Lvl(I)
                        .Hidden = False
                        .Width = 80
                    Else
                        .Hidden = True
                    End If
                End With
            Next
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            .Columns("ITEM_COST_STD").Header.Fixed = True
        End With

        Dim icol_QTY As Integer = 0
        Dim icol_CST As Integer = 0

        For Each C As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "CON", "RTV", "PHY", "ON_HAND", "ONPO", "OPEN", "PICK", "COMM"}
            With grdASTEXPT1.DisplayLayout.Bands(0)
                For Each CTYP As String In New String() {"QTY", "CST"}

                    With .Columns("WHSE_" & CTYP & "_" & C)
                        If icol_QTY = 0 And CTYP = "QTY" And C = "COMM" Then icol_QTY = .Header.VisiblePosition
                        If CTYP = "CST" And C = "COMM" Then icol_CST = .Header.VisiblePosition
                        .Width = 80
                        .Hidden = False
                        .Format = "#,##0"
                        .Header.Caption = CTYP & " " & C
                        If CTYP = "QTY" Then
                            .Header.Appearance.BackColor2 = Color.LightBlue
                        Else
                            .Header.Appearance.BackColor2 = Color.LightGreen
                        End If
                    End With
                    Create_Summary(grdASTEXPT1, "WHSE_" & CTYP & "_" & C)
                Next
            End With
        Next

        With grdASTEXPT1.DisplayLayout.Bands(0).Columns("QTY_AVA")
            .Width = 80
            .Hidden = False
            .Format = "#,##0"
            .Header.Caption = "Qty Ava"
            .Header.Appearance.BackColor2 = Color.Orange
            .Header.VisiblePosition = icol_QTY + 1
        End With
        Create_Summary(grdASTEXPT1, "QTY_AVA")

        With grdASTEXPT1.DisplayLayout.Bands(0).Columns("CST_AVA")
            .Width = 80
            .Hidden = False
            .Format = "#,##0"
            .Header.Caption = "Cst Ava"
            .Header.Appearance.BackColor2 = Color.Orange
            .Header.VisiblePosition = icol_CST
        End With
        Create_Summary(grdASTEXPT1, "CST_AVA")

        Sort_grdColumns(grdASTEXPT1, "G1, G2, G3, G4, G5, G6, G7, G8, G9, ITEM_CODE")
        UltraTabControl1.Tabs("Data Exports").Visible = (optUC.Value = "U")

    End Sub
    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sql As String = String.Empty

        EnforceConstraints(False)
        CDCaption = ""
        optTD = optCOSTDM.Value
        If optTD = "D" Then
            CDCaption = "Direct, "
        End If

        If optCOSTTL.Value = "T" Then
            optTL = "T"
            CDCaption = CDCaption & "Total Unit Costs"
        Else
            optTL = "L"
            CDCaption = CDCaption & "Landed Unit Costs"
        End If

        If optUC.Value = "U" Then
            CDCaption = "Showing Units"
        Else
            CDCaption = "Showing Extended " & CDCaption
        End If

        ASCMAIN1.Progress("Now Processing", "")

        sql = "Select * from ICTSTAT1 where OPS_YYYYPP between '" & RYP & "' and '" & RYP1 & "'"
        Dim TT_ICTSTAT1 As String = ASCMAIN1.Temp_Table(sql)
        sql = "ALTER TABLE " & TT_ICTSTAT1 & " ADD PRIMARY KEY (OPS_YYYYPP, ITEM_CODE, WHSE_CODE)"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Insert into " & TT_ICTSTAT1 & " (OPS_YYYYPP, ITEM_CODE, WHSE_CODE) "
        sql &= " (Select '" & RYP1 & "' OPS_YYYYPP, ITEM_CODE, WHSE_CODE"
        If RYP1 <> ASCMAIN1.CYP Then
            sql &= " from ICTSTAT5 "
            sql &= " where OPS_YYYYPP = '" & RYP1 & "'"
        Else
            sql &= " from ICTSTAT2 "
        End If
        sql &= " minus SELECT OPS_YYYYPP, ITEM_CODE, WHSE_CODE from " & TT_ICTSTAT1 & ")"
        ASCDATA1.ExecuteSQL(sql)

        Dim sqlWhse As String = SQL_in("WHSE_CODE", "A.WHSE_CODE")
        If chkALLITEMS.Checked AndAlso sqlWhse <> String.Empty Then
            sql = "Insert into " & TT_ICTSTAT1 & " (OPS_YYYYPP, ITEM_CODE, WHSE_CODE) "
            sql &= " (Select '" & RYP1 & "' OPS_YYYYPP, ICTITEM1.ITEM_CODE, ICTWHSE1.WHSE_CODE"
            sql &= "  from ICTITEM1,ICTWHSE1"
            sql &= "  where ICTWHSE1.WHSE_CODE " & sqlWhse
            If chkActive.Checked Then
                sql &= "  and ICTITEM1.ITEM_STATUS = 'A'"
            End If
            sql &= " minus SELECT OPS_YYYYPP, ITEM_CODE, WHSE_CODE from " & TT_ICTSTAT1 & ")"
            ASCDATA1.ExecuteSQL(sql)
        End If

        Dim sqlICTCOSTA As String = String.Empty
        Dim C As String = ""
        Select Case optUC.Value
            Case "U"
                sqlICTCOSTA = "ICTCOSTA.ITEM_COST_TOTAL "
            Case "C"
                Select Case optTD & optTL
                    Case "TT"
                        sqlICTCOSTA = "ICTCOSTA.ITEM_COST_TOTAL "
                    Case "TL"
                        sqlICTCOSTA = "(ICTCOSTA.ITEM_COST_VCOST + ICTCOSTA.ITEM_COST_MATLS + ICTCOSTA.ITEM_COST_LANDG + ICTCOSTA.ITEM_COST_LANDGI) "
                    Case "DT"
                        sqlICTCOSTA = "(ICTCOSTA.ITEM_COST_VCOST + ICTCOSTA.ITEM_COST_LANDG + ICTCOSTA.ITEM_COST_TOOLG + ICTCOSTA.ITEM_COST_OVRHD) "
                    Case "DL"
                        sqlICTCOSTA = "(ICTCOSTA.ITEM_COST_VCOST + ICTCOSTA.ITEM_COST_LANDG) "
                End Select
                C = " * " & sqlICTCOSTA
        End Select

        MyBase.Get_SQL("*", TT_ICTSTAT1)
        'prob should have cols for qty and separate cols for cst
        sql = "Select " & sql_SELECT_cols & IIf(COLUMN_NAMEs.Contains("WHSE_CODE"), "", ", ICTISTA1.WHSE_CODE") & ", ICTISTA1.OPS_YYYYPP, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM"
        sql &= ", " & sqlICTCOSTA & " ITEM_COST_STD "
        sql &= ", ICTISTA1.WHSE_QTY_BEG " & C & " WHSE_QTY_BEG, ICTISTA1.WHSE_QTY_SHP " & C & " WHSE_QTY_SHP, ICTISTA1.WHSE_QTY_RTN " & C & " WHSE_QTY_RTN"
        sql &= ", ICTISTA1.WHSE_QTY_REC " & C & " WHSE_QTY_REC, ICTISTA1.WHSE_QTY_ADJ " & C & " WHSE_QTY_ADJ, ICTISTA1.WHSE_QTY_XFR " & C & " WHSE_QTY_XFR"
        sql &= ", ICTISTA1.WHSE_QTY_CON " & C & " WHSE_QTY_CON, ICTISTA1.WHSE_QTY_RTV " & C & " WHSE_QTY_RTV, ICTISTA1.WHSE_QTY_PHY " & C & " WHSE_QTY_PHY"
        sql &= ", B.WHSE_QTY_ON_HAND " & C & " WHSE_QTY_ON_HAND, B.WHSE_QTY_ONPO " & C & " WHSE_QTY_ONPO"
        sql &= ", B.WHSE_QTY_OPEN " & C & " WHSE_QTY_OPEN, B.WHSE_QTY_PICK " & C & " WHSE_QTY_PICK"
        sql &= ", B.WHSE_QTY_COMM " & C & " WHSE_QTY_COMM"
        If RYP1 <> ASCMAIN1.CYP Then
            sql &= " from ICTSTAT5 B, " & TT_ICTSTAT1 & " ICTISTA1, ICTITEM1, ICTCOSTA"
        Else
            sql &= " from ICTSTAT2 B, " & TT_ICTSTAT1 & " ICTISTA1, ICTITEM1, ICTCOSTA"
        End If

        sql_TABLE_NAMEs = sql_TABLE_NAMEs.Replace(",ICTITEM1", "")
        sql &= sql_TABLE_NAMEs

        sql &= " where B.ITEM_CODE (+) = ICTISTA1.ITEM_CODE"
        sql &= "   and B.WHSE_CODE (+) = ICTISTA1.WHSE_CODE"
        sql &= "   and ICTITEM1.ITEM_CODE = ICTISTA1.ITEM_CODE"
        sql &= "   and ICTCOSTA.ITEM_CODE (+) = ICTISTA1.ITEM_CODE"
        sql &= "   and ICTCOSTA.OPS_YYYYPP (+) = ICTISTA1.OPS_YYYYPP"
        If chkActive.Checked Then
            sql &= "   and ICTITEM1.ITEM_STATUS = 'A'"
        End If

        If RYP1 <> ASCMAIN1.CYP Then
            sql &= " and B.OPS_YYYYPP (+) = '" & RYP1 & "'"
        End If

        sql &= sql_JOIN
        sql &= sql_WHERE
        Dim TT_DATA As String = ASCMAIN1.Temp_Table(sql)

        sql = "Select * From " & TT_DATA

        Dim getColumns As String = String.Empty

        For Each field As String In sql_SELECT_cols.Split(",")
            field = field.Trim
            If Not field.StartsWith("'x'") Then
                getColumns &= field.ToUpper.Split(" AS ")(2).Trim
            Else
                getColumns &= field.ToUpper.Replace("'X' AS", "").Trim
            End If
            getColumns &= ","
        Next

        'sql_SELECT_cols
        sql = "Select " & getColumns
        sql &= " ITEM_CODE, SUM(WHSE_QTY_BEG), SUM(WHSE_QTY_SHP), SUM(WHSE_QTY_RTN) " 'WHSE_CODE,
        sql &= ", SUM(WHSE_QTY_REC), SUM(WHSE_QTY_ADJ), SUM(WHSE_QTY_XFR) "
        sql &= ", SUM(WHSE_QTY_CON), SUM(WHSE_QTY_RTV), SUM(WHSE_QTY_PHY) "
        sql &= ", SUM(WHSE_QTY_ON_HAND), SUM(WHSE_QTY_ONPO) "
        sql &= ", SUM(WHSE_QTY_OPEN), SUM(WHSE_QTY_PICK)"
        sql &= ", SUM(WHSE_QTY_COMM)"
        sql &= " from " & TT_DATA
        sql &= " group by " & getColumns & "ITEM_CODE" ' , WHSE_CODE"
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        sql = "Select ITEM_CODE, ITEM_DESC, ITEM_UOM, ITEM_COST_STD" _
            & " from ICTITEM1 where ITEM_CODE in (Select ITEM_CODE from " & TT_DATA & ")"
        With dst
            If Not .Tables.Contains("ICTITEM1") Then
                Create_TDA(.Tables.Add, "ICTITEM1", sql, 0, False, "", 1)
            Else
                dst.Tables("ICTITEM1").Clear()
            End If
            Fill_Records("ICTITEM1", String.Empty, True, sql)
        End With

        EnforceConstraints(True)

    End Sub
End Class